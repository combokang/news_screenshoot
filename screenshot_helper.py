import asyncio
from pyppeteer import launch, errors
import time
import excelhandler
import os


print("開始執行截圖...")


async def main(url_dic):
    # 以excel名稱命名建立資料夾，若已存在同名資料夾則略過
    new_path = f"scrnshots-{excelhandler.file_name}"
    if not os.path.exists(new_path):
        os.mkdir(new_path)
    # 爬蟲並截圖
    for key, value in url_dic.items():
        try:
            # 檢查截圖是否已存在，存在的話直接略過截圖步驟
            image_path = f"{new_path}/{key}.png"
            if os.path.isfile(image_path):
                print(f"略過步驟：編號: {key} 的截圖已存在")
            else:
                browser = await launch({"dumpio": True})
                context = await browser.createIncognitoBrowserContext()  # 開啟無痕模式
                page = await context.newPage()
                try:
                    await page.goto(value, timeout=60000)
                except errors.TimeoutError as e:
                    print("連線逾時")

                # 如果是自由時報新聞、地產天下，自動按下取消訂閱，再捲動到新聞底端以擷取到lazy load圖片
                if "ltn.com.tw" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            document.getElementsByClassName('softPush_refuse')[0].dispatchEvent(new MouseEvent('click', {{
                                bubbles: true,
                                cancelable: true,
                                view: window
                            }}));
                        }}""")
                    except:
                        pass
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementsByClassName('text boxTitle')[0]
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是蘋果地產新聞，讓畫面捲動到新聞底端以擷取到lazy load圖片
                elif "tw.feature.appledaily.com" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementsByClassName('tds_article')[0]
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是蘋果新聞，讓畫面捲動到新聞底端以擷取到lazy load圖片
                elif "tw.appledaily.com" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementById('articleBody')
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是新浪台灣，讓畫面捲動到新聞底端以擷取到lazy load圖片
                elif "tw.appledaily.com" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementById('artContent')
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是新浪新聞，讓畫面捲動到新聞底端以擷取到lazy load圖片
                elif "news.sina.com.tw" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementById('article_content')
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是UDN新聞，捲動到新聞底端以擷取到lazy load圖片
                elif "udn.com" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            element = document.getElementsByClassName('article-content')[0]
                            window.scrollBy({{
                                top: element.scrollHeight, 
                                behavior: "smooth"
                            }});
                        }}""")
                    except:
                        pass
                # 如果是卡優新聞，嘗試縮小廣告
                elif "cardu.com.tw" in value:
                    try:
                        await page.evaluate(f"""() => {{
                            javascript:carzy_showSmallLayer();
                        }}""")
                    except:
                        pass

                time.sleep(2)   # 等待2秒(部分網站如新浪需要約至少2秒才能讀取到圖片)
                await page.screenshot(path=image_path, fullPage=True)
                print(f"編號: {key} 已完成截圖")
                await browser.close()
        except Exception as e:
            print(e)
    print("程式已運行完畢\n")

while True:
    asyncio.get_event_loop().run_until_complete(main(excelhandler.url_dic))
    print("若有新聞需要重新截圖，請將該新聞的舊截圖移出資料夾或刪除")
    while True:
        flag = input("請輸入1 + Enter以重新執行截圖(已有截圖的新聞將被略過)，或直接輸入Enter以退出程式：")
        if flag == "1" or flag == "":
            break
    if flag != "1":
        break
