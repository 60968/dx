import requests
import time
import csv
import os
import json

# 获取B站视频下的评论及子评论

COOKIE = input(
    "按回车继续（如有报错，请输入Cookie）： ") or "SESSDATA=ff87907b%2C1779676271%2Cc78b4%2Ab1CjB13cjL5IGPPI1JzRmjGbPdAPyE5_U4YCIMuQCG5zddUMqsRchN5UoDnAJmwkc5Pi0SVjZCblZZVkt6d3gyYk45YWlIYm9LdVJEUHEwLXdzdllvcDdhRWw4Y3V5WjkzVWFRaF84TmtvR1k0aHZPNVVVWXBBRUpHdGxIb2xNb1dlYXhaYlRQbkJ3IIEC; bili_jct=baa3eda6b2140a2bdcc084d3f0b7bde1"


def get_video_oid(bvid):
    """将BV号转换为oid（视频数字ID）"""
    url = f"https://api.bilibili.com/x/web-interface/view?bvid={bvid}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Cookie": COOKIE,
        "Referer": "https://www.bilibili.com/"
    }

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()
        if data.get("code") == 0:
            return data["data"]["aid"]
        else:
            print(f"获取oid失败: {data.get('message', '未知错误')}")
            return None
    except Exception as e:
        print(f"获取oid时出错: {str(e)}")
        return None


def get_comments(oid, page=1, max_comments=10000):
    """获取主评论（包含子评论）- 完全修复NoneType错误"""
    url = "https://api.bilibili.com/x/v2/reply"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Cookie": COOKIE,
        "Referer": f"https://www.bilibili.com/video/{oid}"
    }

    params = {
        "type": "1",
        "oid": oid,
        "pn": page,
        "ps": 20  # 每页20条
    }

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        # 安全获取data字段
        data_content = data.get("data", {})
        replies = data_content.get("replies", [])

        comments = []
        for item in replies:
            comment_data = {
                "user": item["member"]["uname"],
                "content": item["content"]["message"],
                "like_count": item["like"],
                "time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(item["ctime"])),
                "rpid": item["rpid"]
            }
            comments.append(comment_data)

            # 获取子评论（只在有子评论时才请求）
            if item["count"] > 0:
                sub_comments = get_sub_comments(oid, item["rpid"])
                comment_data["sub_comments"] = sub_comments
            else:
                comment_data["sub_comments"] = []

        return comments
    except Exception as e:
        print(f"获取评论时出错: {str(e)}")
        return []


def get_sub_comments(oid, root_rpid):
    """获取子评论 - 修复NoneType错误"""
    url = "https://api.bilibili.com/x/v2/reply/reply"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Cookie": COOKIE,
        "Referer": f"https://www.bilibili.com/video/{oid}"
    }

    params = {
        "type": "1",
        "oid": oid,
        "root": root_rpid,
        "pn": 1,
        "ps": 20  # 每页20条
    }

    try:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        # 安全获取data字段
        data_content = data.get("data", {})
        sub_replies = data_content.get("replies", [])

        sub_comments = []
        for item in sub_replies:
            sub_comments.append({
                "user": item["member"]["uname"],
                "content": item["content"]["message"],
                "like_count": item["like"],
                "time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(item["ctime"]))
            })

        return sub_comments
    except Exception as e:
        print(f"获取子评论时出错: {str(e)}")
        return []


def save_to_csv(comments, filename="bilibili_comments.csv"):
    """将评论保存到CSV文件"""
    file_exists = os.path.isfile(filename)

    with open(filename, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)

        if not file_exists:
            writer.writerow(["用户", "评论内容", "点赞数", "时间", "子评论数", "子评论"])

        for comment in comments:
            # 确保sub_comments是列表
            if not isinstance(comment["sub_comments"], list):
                comment["sub_comments"] = []

            sub_comments_str = json.dumps(comment["sub_comments"], ensure_ascii=False)
            writer.writerow([
                comment["user"],
                comment["content"],
                comment["like_count"],
                comment["time"],
                len(comment["sub_comments"]),
                sub_comments_str
            ])


def main():
    # 设置BV号缺省值
    bvid = input("输入视频BV号（如: BV1U9UbBUEVv）: ") or "BV1U9UbBUEVv"

    # 获取视频oid
    oid = get_video_oid(bvid)
    if not oid:
        print("无法获取视频ID，请检查BV号或Cookie")
        return

    print("爬取速度控制：每100条暂停5秒，每1000条暂停30秒")

    total_comments = 0
    page = 1
    max_comments = 10000

    while total_comments < max_comments:
        comments = get_comments(oid, page)

        if not comments:
            break

        save_to_csv(comments)
        total_comments += len(comments)
        print(f"已获取 {len(comments)} 条评论，累计 {total_comments} 条")

        # 速度控制
        if total_comments % 100 == 0 and total_comments > 0:
            print(f"已获取 {total_comments} 条，暂停5秒...")
            time.sleep(5)

        if total_comments % 1000 == 0 and total_comments > 0:
            print(f"已获取 {total_comments} 条，暂停30秒...")
            time.sleep(30)

        page += 1

    print(f"爬取完成！共获取 {total_comments} 条评论，已保存到 bilibili_comments.csv")


if __name__ == "__main__":
    main()