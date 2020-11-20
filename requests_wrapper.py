import requests
from lxml import etree

from proxy import get_random_ua


class RequestsWrapper(object):
    proxies = None
    user_agent = None
    request = requests
    cookies = None
    HTTPError = requests.exceptions.HTTPError

    def __init__(self, use_session=False, use_proxy=True, proxies=proxies):
        self.user_agent = get_random_ua()

        if use_proxy is True:
            self.proxies = proxies

        if use_session:
            self.request = requests.Session()

    def get(self, url, **kwargs):
        if 'headers' in kwargs:
            kwargs['headers']['User-Agent'] = self.user_agent
        else:
            kwargs['headers'] = {'User-Agent': self.user_agent}

        return self.request.get(
            url=url,
            **kwargs
        )

    def post(self, url, **kwargs):
        if 'headers' in kwargs:
            kwargs['headers']['User-Agent'] = self.user_agent
        else:
            kwargs['headers'] = {'User-Agent': self.user_agent}

        return self.request.post(
            url=url,
            **kwargs
        )

    def get_html_page(self, url, method='get', encoding=None, tree=None, **kwargs):
        if method == 'get':
            web_page = self.get(url, **kwargs)
        elif method == 'post':
            web_page = self.post(url, **kwargs)
        else:
            web_page = None
        web_page.raise_for_status()
        raw_html = web_page.content
        self.cookies = web_page.cookies
        if encoding:
            raw_html = raw_html.decode(encoding)
        if tree:
            html = etree.HTML(raw_html)
        else:
            html = raw_html
        return html

    def get_json_page(self, url, method='get', **kwargs):
        if method == 'get':
            web_page = self.get(url, **kwargs)
        elif method == 'post':
            web_page = self.post(url, **kwargs)
        else:
            web_page = None
        web_page.raise_for_status()
        json = web_page.json()
        return json

    def set_cookie(self, domain, name, value):
        cookie_obj = requests.cookies.create_cookie(domain=domain, name=name, value=value)
        self.request.cookies.set_cookie(cookie_obj)
