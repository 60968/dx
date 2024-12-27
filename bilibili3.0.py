import requests
import time
from datetime import datetime
from docx import Document

def get_comments(api_url, params, headers, retries=3, backoff_factor=1):
    for attempt in range(retries):
        response = requests.get(api_url, params=params, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Attempt {attempt + 1} failed with status code: {response.status_code}")
            print(f"Response content: {response.text}")  # 打印响应内容以便调试
            if attempt < retries - 1:
                time.sleep(backoff_factor * (2 ** attempt))  # 指数退避
            else:
                print(f"Failed to retrieve comments after {retries} attempts.")
                return None

def fetch_all_comments(comments_data, indent=0):
    comments = comments_data.get('data', {}).get('replies', [])
    all_comments = []

    for comment in comments:
        comment_time = datetime.fromtimestamp(comment['ctime']).strftime('%Y-%m-%d %H:%M:%S')
        comment_text = " " * indent + f"[{comment_time}] {comment['content']['message']}"
        all_comments.append(comment_text)

        # 递归获取子评论
        if 'replies' in comment and comment['replies']:
            sub_comments = fetch_all_comments({'data': {'replies': comment['replies']}}, indent + 4)
            all_comments.extend(sub_comments)

    return all_comments

def save_comments_to_docx(comments, filename="主评论及回复.docx"):
    doc = Document()
    for comment in comments:
        doc.add_paragraph(comment)
    doc.save(filename)
    print(f"已保存为文件： {filename}")

def main():
    default_video_id = "BV1TUkCYEEYB"
    default_cookie = "buvid3=C9C0F219-B6EE-2079-CEAA-5D07D550458484128infoc; b_nut=1735263584; b_lsid=BAEDBA25_19405C500DC; bsource=search_bing; _uuid=5ED96B62-BFA2-9AAA-F357-6EEB2BE4735E84487infoc; buvid_fp=3fd299b7faf0b10ac884967013ac70cb; buvid4=A813E979-38E2-A566-686A-8AD4FAE2F8DC84796-024122701-xRHa1wK8bEeayBE1FOiHwQ%3D%3D; enable_web_push=DISABLE; home_feed_column=5; browser_resolution=1528-740; bili_ticket=eyJhbGciOiJIUzI1NiIsImtpZCI6InMwMyIsInR5cCI6IkpXVCJ9.eyJleHAiOjE3MzU1MjI3OTAsImlhdCI6MTczNTI2MzUzMCwicGx0IjotMX0.2wQTtjepmbKnmIz5dDinln3C8z4tx6kACCSsHLWRJkc; bili_ticket_expires=1735522730; rpdid=|(Yuk)R~)Yl0J'u~JlkYJlmJ; SESSDATA=8c20b5db%2C1750815657%2Cef167%2Ac2CjDkEmY6foxg1oRQ6Xd74t-3YnuUz0hQaLsVapXUrzWHBmlfohdmJf8Sz3nlLL6vVyUSVk5wZHpkWmpaTW0tcFhvVlA4ZVZSaHJJaGs5bGdSVTgzYlZ0dF9CaGFkOFdiYjNGMHNCamdGX1phUTlxU05vdHl5emNWNWZjVmVxOUl4ZVZkaS1xaFlBIIEC; bili_jct=3ea000a9471458a5d7a9ca1138498e11; DedeUserID=452770547; DedeUserID__ckMd5=32759f0f81b4dabc; CURRENT_FNVAL=16; CURRENT_QUALITY=80; sid=50nm0tbf; bp_t_offset_452770547=1015449428347060224"

    video_id = input(f"Enter video_id: ").strip() or default_video_id
    cookie = input(f"Enter Cookie: ").strip() or default_cookie
    print(f"程序正在运行……")

    api_url = "https://api.bilibili.com/x/v2/reply"
    headers = {
        'Cookie': cookie,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Safari/537.36',
        'Referer': 'https://www.bilibili.com/',
        'Origin': 'https://www.bilibili.com',
        'Accept': 'application/json'
    }

    page_number = 1
    all_comments = []

    try:
        while True:
            params = {
                "oid": video_id,
                "pn": page_number,  # 页码
                "type": 1  # 类型，1为视频评论
            }

            comments_data = get_comments(api_url, params, headers)
            if comments_data:
                fetched_comments = fetch_all_comments(comments_data)
                if not fetched_comments:
                    break  # 没有更多评论，退出循环
                all_comments.extend(fetched_comments)
                page_number += 1
            else:
                break  # 请求失败，退出循环

        save_comments_to_docx(all_comments)
        print("评论已保存，此窗口5秒后自动关闭")
        time.sleep(5)

    except Exception as e:
        with open("Error.txt", "w", encoding="utf-8") as error_file:
            error_file.write(f"Error occurred: {str(e)}\n")
        print(f"An error occurred: {str(e)}")
        print("Error details saved to Error.txt")

if __name__ == "__main__":
    main()
