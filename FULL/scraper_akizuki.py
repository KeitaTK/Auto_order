"""
秋月電子用スクレイパー
"""
from typing import Optional, Dict, Any
import time
import re
import json
from urllib.parse import urlparse
from bs4 import BeautifulSoup
from scraper_base import ScraperBase


class AkizukiScraper(ScraperBase):
    """秋月電子専用スクレイパー"""
    
    def get_site_name(self) -> str:
        return "秋月電子通商"
    
    def is_valid_url(self, url: str) -> bool:
        """秋月電子のURLかどうか判定"""
        try:
            p = urlparse(url)
            host = (p.netloc or '').lower()
            return 'akizukidenshi.com' in host
        except Exception:
            return False
    
    def fetch_product_data(self, url: str) -> Optional[Dict[str, Any]]:
        """秋月電子の商品ページから商品情報を取得"""
        html = self._fetch_page(url)
        if not html:
            return None
        
        soup = BeautifulSoup(html, 'html.parser')
        
        # Extract fields
        name = self._extract_name(soup)
        model = self._extract_model(soup)
        item_code = self._extract_item_code(soup, url)
        price_ex, price_in = self._extract_prices(soup)
        
        # Fallback: JSON-LD price/name if missing
        if (price_ex is None or price_in is None) or not name:
            jl_name, jl_price = self._extract_jsonld(soup)
            if not name and jl_name:
                name = jl_name
            if price_in is None and jl_price is not None:
                price_in = jl_price
        
        # Normalize to int
        price_ex = self._to_int(price_ex)
        price_in = self._to_int(price_in)
        
        # Derive missing price using 10% tax
        if price_in is None and price_ex is not None:
            price_in = int(round(price_ex * 1.1))
        if price_ex is None and price_in is not None:
            price_ex = int(round(price_in / 1.1))
        
        if not any([name, model, item_code, price_ex, price_in]):
            return None
        
        return {
            'supplier': '秋月電子通商',
            'item_code': item_code or '',
            'name': name or '',
            'model': model or '',
            'price_excl_tax': price_ex if price_ex is not None else 0,
            'price_incl_tax': price_in if price_in is not None else 0,
            'url': url,
        }
    
    def _fetch_page(self, url: str, retries: int = 3, timeout: int = 20):
        """ページを取得"""
        last_err = None
        for i in range(retries):
            try:
                r = self.session.get(url, timeout=timeout)
                if r.status_code == 200 and 'text/html' in r.headers.get('Content-Type', '').lower():
                    text = r.text
                    if any(x in text for x in ['アクセスが集中', 'メンテナンス', 'ただいま処理中']):
                        time.sleep(1.2 + i * 0.5)
                        continue
                    return text
                else:
                    last_err = f'HTTP {r.status_code}'
            except Exception as e:
                last_err = str(e)
            time.sleep(1.0 + i * 0.5)
        return None
    
    def _extract_name(self, soup: BeautifulSoup):
        """商品名を抽出"""
        # Priority: h1 product name
        el = soup.select_one('h1.h1-goods-name')
        if el and el.get_text(strip=True):
            return el.get_text(strip=True)
        
        og = soup.find('meta', property='og:title')
        if og and og.get('content'):
            return og['content'].strip()
        
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
            title = re.sub(r'\s*[｜|:].*$', '', title)
            return title
        return None
    
    def _extract_model(self, soup: BeautifulSoup):
        """型番を抽出"""
        # Direct id
        dd = soup.select_one('dd#spec_number')
        if dd and dd.get_text(strip=True):
            return dd.get_text(strip=True)
        
        # dt "型番" -> sibling dd
        for dt in soup.select('dt, th'):
            label = dt.get_text(strip=True)
            if '型番' in label or '型式' in label or '品番' in label:
                sib = None
                if dt.name == 'dt':
                    sib = dt.find_next_sibling('dd')
                else:
                    sib = dt.find_next_sibling('td')
                if sib and sib.get_text(strip=True):
                    return sib.get_text(strip=True)
        
        # Inline fallback
        text = soup.get_text('\n', strip=True)
        m = re.search(r'(?:型番|型式|品番)\s*[:：]\s*([^\s　|｜\n]+)', text)
        if m:
            return m.group(1)
        return None
    
    def _extract_item_code(self, soup: BeautifulSoup, url: str):
        """商品コードを抽出"""
        # DOM first
        dd = soup.select_one('dd#spec_goods')
        if dd and dd.get_text(strip=True):
            return dd.get_text(strip=True)
        
        for dt in soup.select('dt, th'):
            if '販売コード' in dt.get_text(strip=True) or '商品コード' in dt.get_text(strip=True):
                sib = None
                if dt.name == 'dt':
                    sib = dt.find_next_sibling('dd')
                else:
                    sib = dt.find_next_sibling('td')
                if sib and sib.get_text(strip=True):
                    return sib.get_text(strip=True)
        
        # URL pattern: /catalog/g/g123456/
        try:
            path = urlparse(url).path
            m = re.search(r'/catalog/g/g(\d+)/?', path)
            if m:
                return m.group(1)
        except Exception:
            pass
        
        return None
    
    def _extract_prices(self, soup: BeautifulSoup):
        """価格を抽出（税別、税込）"""
        price_incl = None
        price_excl = None
        
        # Explicit elements
        incl_el = soup.select_one('.block-goods-price--price')
        if incl_el:
            m = re.search(r'￥\s*([0-9,]+)', incl_el.get_text(' ', strip=True))
            if m:
                price_incl = self._to_int(m.group(1))
        
        excl_el = soup.select_one('.block-goods-price--net-price')
        if excl_el:
            m = re.search(r'￥\s*([0-9,]+)', excl_el.get_text(' ', strip=True))
            if m:
                price_excl = self._to_int(m.group(1))
        
        # Context scan with 税込/税抜 hints
        if price_incl is None or price_excl is None:
            text = soup.get_text('\n', strip=True)
            for m in re.finditer(r'￥\s*([0-9,]+)', text):
                val = self._to_int(m.group(1))
                if val is None:
                    continue
                s = max(0, m.start() - 15)
                e = min(len(text), m.end() + 15)
                win = text[s:e]
                if price_incl is None and ('税込' in win or '(税込' in win or '（税込' in win):
                    price_incl = val
                if price_excl is None and ('税抜' in win or '税別' in win):
                    price_excl = val
        
        return price_excl, price_incl
    
    def _extract_jsonld(self, soup: BeautifulSoup):
        """JSON-LDから商品情報を抽出"""
        try:
            for sc in soup.find_all('script', {'type': 'application/ld+json'}):
                content = sc.string or sc.get_text()
                if not content:
                    continue
                data = json.loads(content)
                # Handle list or single object
                if isinstance(data, list):
                    candidates = data
                else:
                    candidates = [data]
                for obj in candidates:
                    if isinstance(obj, dict) and obj.get('@type') in ('Product', 'product', 'Offer'):
                        name = obj.get('name')
                        price = None
                        if 'offers' in obj and isinstance(obj['offers'], dict):
                            price = obj['offers'].get('price')
                        elif 'price' in obj:
                            price = obj.get('price')
                        return (name, self._to_int(price))
        except Exception:
            pass
        return (None, None)
    
    def _to_int(self, v):
        """値をintに変換"""
        if v is None:
            return None
        if isinstance(v, int):
            return v
        try:
            s = str(v)
            s = re.sub(r'[^\d]', '', s)
            if not s:
                return None
            return int(s)
        except Exception:
            return None
