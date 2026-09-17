# -*- coding: utf-8 -*-
"""
微信自动添加好友脚本
读取Excel电话列，通过微信PC端自动搜索手机号添加好友。
前置条件：微信PC客户端已登录并保持在前台运行。
"""
import argparse
import random
import re
import sys
import time
import logging
from pathlib import Path

import pandas as pd
from wxauto import WeChat


def setupLogging(logFile):
    logger = logging.getLogger("wxAdd")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    if logFile:
        fh = logging.FileHandler(logFile, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    return logger


def isValidPhone(phone):
    if pd.isna(phone):
        return False
    s = str(phone).strip()
    s = re.sub(r"[\s\-]", "", s)
    if s.endswith(".0"):
        s = s[:-2]
    return bool(re.fullmatch(r"1[3-9]\d{9}", s))


def normalizePhone(phone):
    s = str(phone).strip()
    s = re.sub(r"[\s\-]", "", s)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def loadExcel(excelPath, phoneCol, remarkCol):
    df = pd.read_excel(excelPath, dtype=str)
    if phoneCol not in df.columns:
        raise ValueError(f"Excel中找不到电话列 '{phoneCol}'，可用列: {list(df.columns)}")
    records = []
    for idx, row in df.iterrows():
        phone = row[phoneCol]
        remark = ""
        if remarkCol and remarkCol in df.columns:
            v = row[remarkCol]
            if not pd.isna(v):
                remark = str(v).strip()
        records.append({"row": idx + 2, "phone": phone, "remark": remark})
    return records


def loadProgress(progressFile):
    if not Path(progressFile).exists():
        return set()
    with open(progressFile, "r", encoding="utf-8") as f:
        return set(line.strip() for line in f if line.strip())


def saveProgress(progressFile, phone):
    with open(progressFile, "a", encoding="utf-8") as f:
        f.write(phone + "\n")


def main():
    parser = argparse.ArgumentParser(description="微信自动添加好友")
    parser.add_argument("--excel", default="企业查询数据(1).xlsx", help="Excel文件路径")
    parser.add_argument("--phone-col", default="电话", help="电话列名")
    parser.add_argument("--remark-col", default="企业名称", help="备注名来源列名，留空则不设备注")
    parser.add_argument("--add-msg", default="您好，我们公司正在主办一场中山数智AI创新交流峰会，专门邀请中山本地的企业主免费参加，学习AI怎么帮企业降本增效，如果你感兴趣可以通过一下进行具体了解详情", help="添加好友验证消息")
    parser.add_argument("--min-interval", type=float, default=8.0, help="每次添加最小间隔秒数")
    parser.add_argument("--max-interval", type=float, default=20.0, help="每次添加最大间隔秒数")
    parser.add_argument("--tags", nargs="*", default=None, help="好友标签列表，如 --tags 客户 中山")
    parser.add_argument("--start-row", type=int, default=0, help="从第几条记录开始(0基,用于断点续传)")
    parser.add_argument("--max-count", type=int, default=0, help="最多添加多少个，0表示不限制")
    parser.add_argument("--progress-file", default="progress.txt", help="已完成手机号记录文件")
    parser.add_argument("--result-file", default="add_result.xlsx", help="结果输出Excel")
    parser.add_argument("--dry-run", action="store_true", help="只读取校验，不实际添加")
    args = parser.parse_args()

    logger = setupLogging("add_friends.log")
    logger.info("=" * 50)
    logger.info(f"Excel: {args.excel}")
    logger.info(f"验证消息: {args.add_msg}")
    logger.info(f"间隔: {args.min_interval}-{args.max_interval}秒")

    records = loadExcel(args.excel, args.phone_col, args.remark_col)
    valid = [r for r in records if isValidPhone(r["phone"])]
    invalid = [r for r in records if not isValidPhone(r["phone"])]
    logger.info(f"总记录 {len(records)} 条，有效手机号 {len(valid)} 条，无效 {len(invalid)} 条")
    if invalid[:5]:
        logger.info(f"无效示例(前5): {[(r['row'], r['phone']) for r in invalid[:5]]}")

    if args.dry_run:
        logger.info("dry-run模式，不实际添加。有效手机号列表(前20):")
        for r in valid[:20]:
            logger.info(f"  行{r['row']} {normalizePhone(r['phone'])} 备注={r['remark']}")
        return

    done = loadProgress(args.progress_file)
    logger.info(f"已完成记录 {len(done)} 个(来自 {args.progress_file})")

    pending = [r for r in valid if normalizePhone(r["phone"]) not in done]
    if args.start_row > 0:
        pending = pending[args.start_row:]
        logger.info(f"从第 {args.start_row} 条开始，剩余 {len(pending)} 条待处理")
    if args.max_count > 0:
        pending = pending[:args.max_count]
        logger.info(f"限制最多添加 {args.max_count} 条")

    logger.info(f"本次待添加 {len(pending)} 条")
    if not pending:
        logger.info("没有待添加的记录，退出。")
        return

    logger.info("正在连接微信PC客户端...")
    try:
        wx = WeChat()
    except Exception as e:
        logger.error(f"连接微信失败: {e}")
        logger.error("请确保微信PC客户端已登录并显示在前台窗口。")
        return
    logger.info("微信连接成功。")

    results = []
    successCount = 0
    failCount = 0

    for i, rec in enumerate(pending, 1):
        phone = normalizePhone(rec["phone"])
        remark = rec["remark"] or None
        logger.info(f"[{i}/{len(pending)}] 行{rec['row']} 手机号 {phone} 备注={remark}")

        status = "unknown"
        errMsg = ""
        try:
            wx.AddNewFriend(
                keywords=phone,
                addmsg=args.add_msg,
                remark=remark,
                tags=args.tags,
            )
            status = "sent"
            successCount += 1
            saveProgress(args.progress_file, phone)
            logger.info(f"  -> 已发送添加请求")
        except Exception as e:
            status = "fail"
            errMsg = str(e)
            failCount += 1
            logger.warning(f"  -> 失败: {errMsg}")

        results.append({
            "行号": rec["row"],
            "手机号": phone,
            "备注": remark or "",
            "状态": status,
            "错误": errMsg,
            "时间": time.strftime("%Y-%m-%d %H:%M:%S"),
        })

        if i < len(pending):
            wait = random.uniform(args.min_interval, args.max_interval)
            logger.info(f"  等待 {wait:.1f} 秒...")
            time.sleep(wait)

    resultDf = pd.DataFrame(results)
    resultDf.to_excel(args.result_file, index=False)
    logger.info("=" * 50)
    logger.info(f"完成。成功发送 {successCount}，失败 {failCount}，结果已保存到 {args.result_file}")


if __name__ == "__main__":
    main()
