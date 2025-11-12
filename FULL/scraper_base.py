"""
スクレイパーの基底クラス
各サイト用のスクレイパーはこのクラスを継承して実装する
"""
from abc import ABC, abstractmethod
from typing import Optional, Dict, Any
import requests


class ScraperBase(ABC):
    """スクレイパー基底クラス"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ja-JP,ja;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1'
        })
    
    @abstractmethod
    def get_site_name(self) -> str:
        """サイト名を返す"""
        pass
    
    @abstractmethod
    def is_valid_url(self, url: str) -> bool:
        """URLがこのサイトのものかどうかを判定"""
        pass
    
    @abstractmethod
    def fetch_product_data(self, url: str) -> Optional[Dict[str, Any]]:
        """
        商品情報を取得
        
        Returns:
            {
                'supplier': str,      # メーカー/仕入元
                'item_code': str,     # 注文コード/商品コード
                'name': str,          # 商品名
                'model': str,         # 品番/型番
                'price_excl_tax': str or int,  # 単価(税別)
                'price_incl_tax': str or int,  # 税込価格
                'url': str           # URL
            }
            または None (取得失敗時)
        """
        pass
