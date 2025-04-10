import whois
import urllib.request
import urllib.error

# 获取域名的whois信息
# print(whois.whois("www.baidu.com"))


def download(url):
    print("Downloading:", url)
    try:
        html = urllib.request.urlopen(url).read()
    except urllib.error.URLError as e:
        print("Download failed:", e.reason)
        html = None
    return html

if __name__ == '__main__':
    url = 'https://haowallpaper.com/homeView'
    # url = 'https://tools.ietf.org/html/rfc7231#section-6'
    html = download(url)

