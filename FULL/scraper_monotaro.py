"""
モノタロウ用スクレイパー
"""
from typing import Optional, Dict, Any
import time
import re
from bs4 import BeautifulSoup
import requests
from scraper_base import ScraperBase


class MonotaroScraper(ScraperBase):
    """モノタロウ専用スクレイパー"""
    
    def get_site_name(self) -> str:
        return "モノタロウ"
    
    def is_valid_url(self, url: str) -> bool:
        """モノタロウのURLかどうか判定"""
        return url.startswith('https://www.monotaro.com')
    
    def fetch_product_data(self, url: str) -> Optional[Dict[str, Any]]:
        """モノタロウの商品ページから商品情報を取得"""
        try:
            # リトライロジック
            response = None
            for attempt in range(3):
                try:
                    response = self.session.get(url, timeout=15)
                    # ステータスコード確認
                    if response.status_code == 403 or 'ログイン' in response.text:
                        # ボット対策またはログイン要求
                        if attempt < 2:
                            time.sleep(2)  # 待機してリトライ
                            continue
                        else:
                            print(f'ログイン要求またはボット対策: {url}')
                            return None
                    if response.status_code != 200:
                        print(f'HTTPエラー {response.status_code}: {url}')
                        return None
                    break
                except requests.exceptions.Timeout:
                    if attempt < 2:
                        time.sleep(2)
                        continue
                    raise
            
            if response is None:
                print(f'HTTPレスポンスが取得できませんでした: {url}')
                return None
            
            # エンコーディング自動検出
            response.encoding = response.apparent_encoding or 'utf-8'
            soup = BeautifulSoup(response.content, 'html.parser')
            
            # URLからコード抽出
            matched = re.search(r'/p/(\d+)/(\d+)', url)
            item_code = None
            if matched:
                item_code = matched.group(1) + matched.group(2)
            
            # ============ 商品名取得 ============
            product_name = ''
            
            # h1タグから取得（メイン商品名）
            h1_tag = soup.select_one('h1')
            if h1_tag:
                product_name = h1_tag.get_text(strip=True)
            
            # h1が見つからない場合は別の方法
            if not product_name:
                title_tag = soup.select_one('title')
                if title_tag:
                    title_text = title_tag.get_text()
                    # タイトルから最初の部分を抽出（"商品名 | モノタロウ"形式）
                    if '|' in title_text:
                        product_name = title_text.split('|')[0].strip()
                    else:
                        product_name = title_text.strip()
            
            if not product_name:
                print(f'商品名が見つかりません: {url}')
                return None
            
            # ============ 型番取得 ============
            model_number = ''
            
            # 方法1: span.AttributeLabelItem から取得（品番M2.5×16 形式）
            attr_labels = soup.select('span.AttributeLabelItem')
            for attr_label in attr_labels:
                text = attr_label.get_text(strip=True)
                # "品番M2.5×16" から "M2.5×16" を抽出
                if '品番' in text:
                    match = re.search(r'品番(.+)', text)
                    if match:
                        model_number = match.group(1).strip()
                        break
            
            # 方法2: タイトルから型番を抽出（M2.5×16 商品名... 形式）
            if not model_number:
                title_tag = soup.select_one('title')
                if title_tag:
                    title_text = title_tag.get_text(strip=True)
                    # パターン: 型番が先頭にある（M2.5×16のような形式）
                    # 数字、アルファベット、×、-、.などを含む可能性
                    match = re.match(r'^([A-Z0-9\.\-]+[×x][A-Z0-9\.\-]+)\s+', title_text, re.IGNORECASE)
                    if match:
                        model_number = match.group(1)
            
            # 方法3: dl > dt/dd から取得（構造化データ）
            if not model_number:
                dts = soup.find_all('dt')
                for dt in dts:
                    dt_text = dt.get_text(strip=True)
                    if '型番' in dt_text or 'SKU' in dt_text or '品番' in dt_text:
                        dd = dt.find_next('dd')
                        if dd:
                            model_number = dd.get_text(strip=True)
                            break
            
            # 方法4: tableのtdから取得
            if not model_number:
                rows = soup.find_all('tr')
                for row in rows:
                    cells = row.find_all(['th', 'td'])
                    if cells and len(cells) >= 2:
                        header = cells[0].get_text(strip=True)
                        if '型番' in header or '品番' in header:
                            model_number = cells[1].get_text(strip=True)
                            break
            
            # ============ 価格取得 ============
            price_tax_excluded = ''  # 税別価格
            price_tax_included = ''  # 税込価格
            
            # 方法1: 販売価格(税別)と販売価格(税込)を直接取得
            # 税別価格: SellingPrice__Title の次の Price--Lg
            selling_price_title = soup.select_one('.SellingPrice__Title')
            if selling_price_title:
                price_elem = selling_price_title.find_next('span', class_='Price--Lg')
                if price_elem:
                    price_text = price_elem.get_text(strip=True)
                    numbers = re.findall(r'\d+', price_text.replace(',', ''))
                    if numbers:
                        price_tax_excluded = numbers[0]
            
            # 税込価格: 販売価格(税込)を含むReferencePrice
            ref_price_title = soup.find('span', class_='ReferencePrice__Title', text=re.compile(r'販売価格.*税込'))
            if ref_price_title and ref_price_title.parent:
                parent_text = ref_price_title.parent.get_text(strip=True)
                numbers = re.findall(r'[\d,]+', parent_text)
                if numbers:
                    price_tax_included = numbers[-1].replace(',', '')
            
            # 方法2: フォールバック - 複数の価格セレクタを試す（税込価格として扱う）
            if not price_tax_included and not price_tax_excluded:
                price_patterns = [
                    'span[class*="price"]',
                    'span[class*="Price"]',
                    '.p-price',
                    '.productPrice',
                    'div[data-testid*="price"]'
                ]
                
                for pattern in price_patterns:
                    price_elements = soup.select(pattern)
                    for elem in price_elements:
                        price_text = elem.get_text(strip=True)
                        # 数字を抽出
                        numbers = re.findall(r'\d+', price_text.replace(',', ''))
                        if numbers:
                            # 最も大きい数字（通常は価格）を税込として扱う
                            price_tax_included = max(numbers, key=int)
                            break
                    if price_tax_included:
                        break
            
            # 方法3: 価格が見つからない場合、テキスト内を直接探索
            if not price_tax_included and not price_tax_excluded:
                page_text = soup.get_text()
                # 「¥xxx」パターンで検索
                price_match = re.search(r'¥([\d,]+)', page_text)
                if price_match:
                    price_tax_included = price_match.group(1).replace(',', '')
            
            # 税別が取得できて税込がない場合は計算
            if price_tax_excluded and not price_tax_included:
                try:
                    price_tax_included = str(int(float(price_tax_excluded) * 1.1))
                except:
                    pass
            
            # 税込が取得できて税別がない場合は計算
            if price_tax_included and not price_tax_excluded:
                try:
                    price_tax_excluded = str(int(float(price_tax_included) / 1.1))
                except:
                    pass
            
            return {
                'supplier': 'モノタロウ',
                'item_code': item_code or '',
                'name': product_name,
                'model': model_number or '',
                'price_excl_tax': price_tax_excluded or '',
                'price_incl_tax': price_tax_included or '',
                'url': url
            }
        
        except Exception as e:
            print(f'エラー: {str(e)}, URL: {url}')
            return None
