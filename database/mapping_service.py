"""
映射服务模块
提供各种映射查询功能，替代原 Excel VLOOKUP
"""
import sys
from pathlib import Path

_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

import pandas as pd
from typing import Dict, List, Optional, Union, Tuple
from datetime import date, datetime

try:
    from database.db_connection import get_db_manager, DatabaseManager
except ImportError:
    from .db_connection import get_db_manager, DatabaseManager


class MappingService:
    """映射查询服务类"""
    
    def __init__(self, db_manager: Optional[DatabaseManager] = None):
        """
        初始化映射服务
        
        Args:
            db_manager: 数据库管理器（如果不提供则使用全局实例）
        """
        self.db = db_manager or get_db_manager()
    
    # ========================================
    # 产品信息映射
    # ========================================
    
    def map_sku_to_product_id(self, sku_list: List[str]) -> Dict[str, str]:
        """
        SKU 批量映射到商品ID
        
        Args:
            sku_list: SKU 列表
        
        Returns:
            映射字典 {sku: product_id}
        """
        if not sku_list:
            return {}
        
        placeholders = ','.join(['%s'] * len(sku_list))
        sql = f"""
            SELECT sku, product_id 
            FROM product_info 
            WHERE sku IN ({placeholders})
        """
        results = self.db.execute_query(sql, tuple(sku_list))
        return {row['sku']: row['product_id'] for row in results if row['product_id']}
    
    def map_sku_to_category(self, sku_list: List[str], level: int = 2) -> Dict[str, str]:
        """
        SKU 映射到分类
        
        Args:
            sku_list: SKU 列表
            level: 分类级别（2=二级分类，3=三级分类）
        
        Returns:
            映射字典 {sku: category}
        """
        if not sku_list:
            return {}
        
        field = f'category_l{level}'
        placeholders = ','.join(['%s'] * len(sku_list))
        sql = f"""
            SELECT sku, {field} 
            FROM product_info 
            WHERE sku IN ({placeholders})
        """
        results = self.db.execute_query(sql, tuple(sku_list))
        return {row['sku']: row[field] for row in results if row[field]}
    
    def map_sku_to_purchase_price(self, sku_list: List[str]) -> Dict[str, float]:
        """
        SKU 映射到采购价
        
        Args:
            sku_list: SKU 列表
        
        Returns:
            映射字典 {sku: purchase_price}
        """
        if not sku_list:
            return {}
        
        placeholders = ','.join(['%s'] * len(sku_list))
        sql = f"""
            SELECT sku, purchase_price 
            FROM product_info 
            WHERE sku IN ({placeholders})
        """
        results = self.db.execute_query(sql, tuple(sku_list))
        return {row['sku']: float(row['purchase_price']) for row in results if row['purchase_price']}
    
    def map_sku_to_supplier(self, sku_list: List[str]) -> Dict[str, str]:
        """
        SKU 映射到供应商
        
        Args:
            sku_list: SKU 列表
        
        Returns:
            映射字典 {sku: supplier}
        """
        if not sku_list:
            return {}
        
        placeholders = ','.join(['%s'] * len(sku_list))
        sql = f"""
            SELECT sku, supplier 
            FROM product_info 
            WHERE sku IN ({placeholders})
        """
        results = self.db.execute_query(sql, tuple(sku_list))
        return {row['sku']: row['supplier'] for row in results if row['supplier']}
    
    # ========================================
    # 站点映射
    # ========================================
    
    def map_account_to_site_identifier(self, account_site_list: List[Tuple[str, str, str]]) -> Dict[Tuple, str]:
        """
        账号+站点+平台 映射到站点识别码
        
        Args:
            account_site_list: [(account_name, site_code, platform), ...]
        
        Returns:
            映射字典 {(account, site, platform): site_identifier}
        """
        if not account_site_list:
            return {}
        
        mapping = {}
        for account, site, platform in account_site_list:
            sql = """
                SELECT site_identifier 
                FROM site_mapping 
                WHERE account_name = %s AND site_code = %s AND platform = %s
                LIMIT 1
            """
            results = self.db.execute_query(sql, (account, site, platform))
            if results and results[0]['site_identifier']:
                mapping[(account, site, platform)] = results[0]['site_identifier']
        
        return mapping
    
    # ========================================
    # 物流费用映射
    # ========================================
    
    def map_logistics_fee(
        self, 
        sku: str, 
        site_code: str, 
        logistics_type: str,
        currency: str = 'eur',
        effective_date: Optional[date] = None
    ) -> Optional[float]:
        """
        查询物流费用
        
        Args:
            sku: 产品编码
            site_code: 站点代码
            logistics_type: 物流类型（头程/尾程/MF/FBA）
            currency: 币种（eur/rmb/usd）
            effective_date: 生效日期（默认为当前日期）
        
        Returns:
            费用金额（找不到返回 None）
        """
        if effective_date is None:
            effective_date = date.today()
        
        fee_field = f'fee_{currency.lower()}'
        
        sql = f"""
            SELECT {fee_field} 
            FROM logistics_mapping 
            WHERE sku = %s 
              AND site_code = %s 
              AND logistics_type = %s
              AND effective_date <= %s
              AND (expired_date IS NULL OR expired_date > %s)
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, (sku, site_code, logistics_type, effective_date, effective_date))
        
        if results and results[0][fee_field] is not None:
            return float(results[0][fee_field])
        return None
    
    def map_logistics_fee_batch(
        self, 
        sku_site_list: List[Tuple[str, str]], 
        logistics_type: str,
        currency: str = 'eur'
    ) -> Dict[Tuple[str, str], float]:
        """
        批量查询物流费用
        
        Args:
            sku_site_list: [(sku, site_code), ...]
            logistics_type: 物流类型
            currency: 币种
        
        Returns:
            映射字典 {(sku, site): fee}
        """
        mapping = {}
        for sku, site in sku_site_list:
            fee = self.map_logistics_fee(sku, site, logistics_type, currency)
            if fee is not None:
                mapping[(sku, site)] = fee
        return mapping
    
    # ========================================
    # 销售负责人映射
    # ========================================
    
    def map_sales_owner(
        self, 
        sku: Optional[str] = None,
        platform: Optional[str] = None,
        site_code: Optional[str] = None,
        mapping_type: Optional[str] = None
    ) -> Optional[str]:
        """
        查询销售负责人
        
        Args:
            sku: 产品编码（可选）
            platform: 平台（可选）
            site_code: 站点（可选）
            mapping_type: 映射类型（平台/站点/SKU，可选）
        
        Returns:
            销售负责人姓名（找不到返回 None）
        """
        conditions = []
        params = []
        
        if sku:
            conditions.append("sku = %s")
            params.append(sku)
        if platform:
            conditions.append("platform = %s")
            params.append(platform)
        if site_code:
            conditions.append("site_code = %s")
            params.append(site_code)
        if mapping_type:
            conditions.append("mapping_type = %s")
            params.append(mapping_type)
        
        if not conditions:
            return None
        
        where_clause = " AND ".join(conditions)
        sql = f"""
            SELECT sales_owner, sales_manager 
            FROM sales_owner_mapping 
            WHERE {where_clause}
              AND (expired_date IS NULL OR expired_date > CURDATE())
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, tuple(params))
        
        if results and results[0]['sales_owner']:
            return results[0]['sales_owner']
        return None
    
    # ========================================
    # 汇率查询
    # ========================================
    
    def get_exchange_rate(
        self, 
        currency_from: str, 
        currency_to: str = 'EUR',
        effective_date: Optional[date] = None
    ) -> Optional[float]:
        """
        查询汇率
        
        Args:
            currency_from: 源币种
            currency_to: 目标币种（默认 EUR）
            effective_date: 生效日期（默认当前日期）
        
        Returns:
            汇率（找不到返回 None）
        """
        if effective_date is None:
            effective_date = date.today()
        
        sql = """
            SELECT rate, rate_type 
            FROM exchange_rate 
            WHERE currency_from = %s 
              AND currency_to = %s
              AND effective_date <= %s
              AND (expired_date IS NULL OR expired_date > %s)
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, (currency_from.upper(), currency_to.upper(), effective_date, effective_date))
        
        if results:
            return float(results[0]['rate'])
        return None
    
    def convert_currency(
        self, 
        amount: float, 
        currency_from: str, 
        currency_to: str = 'EUR'
    ) -> Optional[float]:
        """
        货币转换
        
        Args:
            amount: 金额
            currency_from: 源币种
            currency_to: 目标币种
        
        Returns:
            转换后的金额
        """
        if currency_from.upper() == currency_to.upper():
            return amount
        
        rate = self.get_exchange_rate(currency_from, currency_to)
        if rate is None:
            return None
        
        # 根据汇率类型计算
        # 例如：RMB → EUR 使用除法（7.3），USD → EUR 使用乘法（0.858）
        sql = """
            SELECT rate_type 
            FROM exchange_rate 
            WHERE currency_from = %s AND currency_to = %s
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, (currency_from.upper(), currency_to.upper()))
        
        if results and results[0]['rate_type'] == 'divide':
            return amount / rate
        else:
            return amount * rate
    
    # ========================================
    # VAT 和佣金费率
    # ========================================
    
    def get_vat_rate(self, platform: str, site_code: str) -> Optional[float]:
        """
        查询 VAT 税率
        
        Args:
            platform: 平台
            site_code: 站点
        
        Returns:
            VAT 税率（如 0.19 表示 19%）
        """
        sql = """
            SELECT vat_rate 
            FROM fee_rate_mapping 
            WHERE platform = %s 
              AND site_code = %s
              AND (expired_date IS NULL OR expired_date > CURDATE())
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, (platform, site_code))
        
        if results and results[0]['vat_rate'] is not None:
            return float(results[0]['vat_rate'])
        return None
    
    def get_commission_rate(self, platform: str, site_code: str) -> Optional[float]:
        """
        查询佣金率
        
        Args:
            platform: 平台
            site_code: 站点
        
        Returns:
            佣金率（如 0.15 表示 15%）
        """
        sql = """
            SELECT commission_rate 
            FROM fee_rate_mapping 
            WHERE platform = %s 
              AND site_code = %s
              AND (expired_date IS NULL OR expired_date > CURDATE())
            ORDER BY effective_date DESC 
            LIMIT 1
        """
        results = self.db.execute_query(sql, (platform, site_code))
        
        if results and results[0]['commission_rate'] is not None:
            return float(results[0]['commission_rate'])
        return None


# 使用示例
if __name__ == "__main__":
    # 创建映射服务实例
    service = MappingService()
    
    print("\n=== 测试映射服务 ===\n")
    
    # 测试汇率查询
    print("1. 查询汇率：")
    rate_rmb_eur = service.get_exchange_rate('RMB', 'EUR')
    print(f"   RMB → EUR: {rate_rmb_eur}")
    
    # 测试货币转换
    print("\n2. 货币转换：")
    amount_eur = service.convert_currency(730, 'RMB', 'EUR')
    print(f"   730 RMB = {amount_eur} EUR")
    
    print("\n✅ 映射服务测试完成！")
