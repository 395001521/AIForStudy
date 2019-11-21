import requests
from bs4 import BeautifulSoup

base_url='https://www.ppzuowen.com/'
user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36\
    (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'

header = {'User_Agent':user_agent}
Title=''

def get_sub_url():
    items = []
    r = requests.get('http://www.ppzuowen.com/book/lvyexianzong/', headers = header)
    r.encoding = 'gbk'
    soup = BeautifulSoup(r.text, 'html.parser')
    
    global Title
    title = soup.select('h2[class="articleH22"]')[0].text
    if title != None:
        Title = title
    
    for sutitle in soup.find_all('a', 'title'):
        suburls = {}
        suburls['link'] = sutitle.get('href')
        items.append(suburls)
    return items

def get_content(url):
    r = requests.get(url, headers = header)
    r.encoding = 'gb2312'
    soup = BeautifulSoup(r.text, 'html.parser')
    contents = soup.find_all(class_ = "articleContent")
    title = '\n'
    title += soup.find(class_ = "articleH2").text + '\n'
    for content in contents:
        with open(Title, "a+", encoding='utf-8') as f:
            f.write(title)
            f.write(str(content.get_text()).replace('<br/>', '\n'))
     
if __name__ == '__main__':
    sub_urls = get_sub_url()
    for i in sub_urls:
        print("Real link:{}".format(base_url + i['link']))
        get_content("{}".format(base_url + i['link']))


