"""
Excel 数据迁移到数据库脚本
将现有的 Excel 映射表批量导入到 MySQL 数据库

使用方法：
    python migrate_excel_to_db.py

注意事项：
    1. 运行前请先执行 schema.sql 创建表结构
    2. 请根据实际 Excel 文件路径修改代码
    3. 首次运行建议先备份原始 Excel 文件
    4. 迁移过程可能需要几分钟，请耐心等待
"""
import pandas as pd
import pymysql
from pathlib import Path
from datetime import date
import sys

_root = Path(__file__).resolve().parent.parent
if str(_root) not in sys.path:
    sys.path.insert(0, str(_root))

from database.db_connection import get_db_manager, DatabaseConfig


class ExcelDataMigrator:
    """Excel 数据迁移器"""
    
    def __init__(self):
        """初始化数据库连接"""
        try:
            self.db = get_db_manager()
            print("✅ 数据库连接成功")
        except Exception as e:
            print(f"❌ 数据库连接失败：{e}")
            print("请检查：")
            print("  1. MySQL 服务是否已启动")
            print("  2. database/db_config.json 配置是否正确")
            print("  3. 数据库 bth-report 是否已创建（与 .env MYSQL_DATABASE 一致）")
            sys.exit(1)
    
    def _get_excel_path(self, filename: str) -> Path:
        """
        获取 Excel 文件路径
        
        Args:
            filename: Excel 文件名
        
        Returns:
            文件完整路径
        """
        # 尝试几个可能的路径
        possible_paths = [
            Path.home() / "Desktop" / filename,
            Path("桌面") / filename,
            Path(f"C:/Users/{Path.home().name}/Desktop") / filename,
            Path(__file__).parent.parent.parent.parent / filename
        ]
        
        for path in possible_paths:
            if path.exists():
                return path
        
        print(f"⚠ 未找到文件：{filename}")
        print("请输入文件的完整路径：")
        manual_path = input().strip().strip('"')
        return Path(manual_path)
    
    def migrate_product_info(self):
        """迁移产品信息库"""
        print("\n" + "=" * 60)
        print("开始迁移：产品信息库")
        print("=" * 60)
        
        try:
            excel_path = self._get_excel_path("产品信息库.xlsx")
            print(f"读取文件：{excel_path}")
            
            # 读取 Excel（跳过前4行，删除前5列）
            df = pd.read_excel(excel_path, sheet_name="产品信息", skiprows=4)
            df = df.iloc[:, 5:]
            
            print(f"读取到 {len(df)} 条记录")
            
            # 准备插入数据
            sql = """
                INSERT INTO product_info 
                (sku, product_id, ean, category_l2, category_l3, supplier, 
                 operation_mode, product_status, purchase_price, declared_price)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    product_id = VALUES(product_id),
                    category_l2 = VALUES(category_l2),
                    category_l3 = VALUES(category_l3),
                    supplier = VALUES(supplier),
                    purchase_price = VALUES(purchase_price),
                    updated_at = CURRENT_TIMESTAMP
            """
            
            data_list = []
            for _, row in df.iterrows():
                # 根据实际 Excel 列名调整
                data_list.append((
                    str(row.get('产品编码', '')).strip(),
                    str(row.get('商品ID', '')).strip() if pd.notna(row.get('商品ID')) else None,
                    str(row.get('EAN', '')).strip() if pd.notna(row.get('EAN')) else None,
                    str(row.get('二级分类', '')).strip() if pd.notna(row.get('二级分类')) else None,
                    str(row.get('三级分类', '')).strip() if pd.notna(row.get('三级分类')) else None,
                    str(row.get('供应商', '')).strip() if pd.notna(row.get('供应商')) else None,
                    str(row.get('运营模式', '')).strip() if pd.notna(row.get('运营模式')) else None,
                    str(row.get('产品状态', '')).strip() if pd.notna(row.get('产品状态')) else None,
                    float(row.get('原始采购价', 0)) if pd.notna(row.get('原始采购价')) else None,
                    float(row.get('申报价格', 0)) if pd.notna(row.get('申报价格')) else None,
                ))
            
            # 批量插入
            affected = self.db.execute_many(sql, data_list)
            print(f"✅ 产品信息库迁移完成：{affected} 条记录")
            
        except FileNotFoundError:
            print("❌ 文件不存在，跳过产品信息库迁移")
        except Exception as e:
            print(f"❌ 产品信息库迁移失败：{e}")
            import traceback
            traceback.print_exc()
    
    def migrate_logistics_mapping(self):
        """迁移物流费用映射表"""
        print("\n" + "=" * 60)
        print("开始迁移：物流费用映射表")
        print("=" * 60)
        
        try:
            # MANO-MF 尾程
            excel_path = self._get_excel_path("MANO-MF尾程.xlsx")
            print(f"读取文件：{excel_path}")
            
            df = pd.read_excel(excel_path, sheet_name=0)
            print(f"读取到 {len(df)} 条记录")
            
            sql = """
                INSERT INTO logistics_mapping 
                (sku, site_code, logistics_type, logistics_category, fee_eur, effective_date)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    fee_eur = VALUES(fee_eur),
                    updated_at = CURRENT_TIMESTAMP
            """
            
            data_list = []
            for _, row in df.iterrows():
                data_list.append((
                    str(row.get('SKU', '')).strip(),
                    str(row.get('站点', '')).strip(),
                    '尾程',
                    'MF',
                    float(row.get('派送费', 0)) if pd.notna(row.get('派送费')) else None,
                    date.today()
                ))
            
            affected = self.db.execute_many(sql, data_list)
            print(f"✅ 物流费用映射表迁移完成：{affected} 条记录")
            
        except FileNotFoundError:
            print("❌ 文件不存在，跳过物流费用映射表迁移")
        except Exception as e:
            print(f"❌ 物流费用映射表迁移失败：{e}")
    
    def migrate_price_mapping(self):
        """迁移定价映射表"""
        print("\n" + "=" * 60)
        print("开始迁移：定价映射表")
        print("=" * 60)
        
        try:
            excel_path = self._get_excel_path("欧洲平台定价表.xlsx")
            print(f"读取文件：{excel_path}")
            
            df = pd.read_excel(excel_path, sheet_name=0, skiprows=2)
            print(f"读取到 {len(df)} 条记录")
            
            sql = """
                INSERT INTO price_mapping 
                (sku, site_code, platform, price_eur, shipping_fee_eur, effective_date)
                VALUES (%s, %s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    price_eur = VALUES(price_eur),
                    shipping_fee_eur = VALUES(shipping_fee_eur),
                    updated_at = CURRENT_TIMESTAMP
            """
            
            data_list = []
            for _, row in df.iterrows():
                data_list.append((
                    str(row.get('产品编码', '')).strip(),
                    str(row.get('站点', '')).strip(),
                    str(row.get('平台', '')).strip() if pd.notna(row.get('平台')) else None,
                    float(row.get('产品单价（EUR）', 0)) if pd.notna(row.get('产品单价（EUR）')) else None,
                    float(row.get('运费回款（EUR）', 0)) if pd.notna(row.get('运费回款（EUR）')) else None,
                    date.today()
                ))
            
            affected = self.db.execute_many(sql, data_list)
            print(f"✅ 定价映射表迁移完成：{affected} 条记录")
            
        except FileNotFoundError:
            print("❌ 文件不存在，跳过定价映射表迁移")
        except Exception as e:
            print(f"❌ 定价映射表迁移失败：{e}")
    
    def migrate_sales_owner_mapping(self):
        """迁移销售负责人映射表"""
        print("\n" + "=" * 60)
        print("开始迁移：销售负责人映射表")
        print("=" * 60)
        
        try:
            # 假设从目标拆解表中读取
            excel_path = self._get_excel_path("26.5月目标拆解及跟进.xlsx")
            print(f"读取文件：{excel_path}")
            
            # 读取多个 sheet（根据实际情况调整）
            sheets = pd.ExcelFile(excel_path).sheet_names
            print(f"找到 {len(sheets)} 个 sheet 页")
            
            all_data = []
            for sheet_name in sheets:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                if 'SKU' in df.columns and '销售负责人' in df.columns:
                    for _, row in df.iterrows():
                        all_data.append((
                            str(row.get('SKU', '')).strip(),
                            str(row.get('销售负责人', '')).strip() if pd.notna(row.get('销售负责人')) else None,
                            str(row.get('销售经理', '')).strip() if pd.notna(row.get('销售经理')) else None,
                        ))
            
            if not all_data:
                print("⚠ 未找到销售负责人数据，跳过")
                return
            
            sql = """
                INSERT INTO sales_owner_mapping 
                (sku, mapping_type, sales_owner, sales_manager, effective_date)
                VALUES (%s, %s, %s, %s, %s)
                ON DUPLICATE KEY UPDATE 
                    sales_owner = VALUES(sales_owner),
                    sales_manager = VALUES(sales_manager),
                    updated_at = CURRENT_TIMESTAMP
            """
            
            data_list = [
                (sku, 'SKU', owner, manager, date.today())
                for sku, owner, manager in all_data
                if owner
            ]
            
            affected = self.db.execute_many(sql, data_list)
            print(f"✅ 销售负责人映射表迁移完成：{affected} 条记录")
            
        except FileNotFoundError:
            print("❌ 文件不存在，跳过销售负责人映射表迁移")
        except Exception as e:
            print(f"❌ 销售负责人映射表迁移失败：{e}")
    
    def verify_data(self):
        """验证迁移数据"""
        print("\n" + "=" * 60)
        print("数据验证")
        print("=" * 60)
        
        tables = [
            'product_info',
            'logistics_mapping',
            'price_mapping',
            'sales_owner_mapping',
            'exchange_rate'
        ]
        
        for table in tables:
            sql = f"SELECT COUNT(*) as count FROM {table}"
            result = self.db.execute_query(sql)
            count = result[0]['count'] if result else 0
            print(f"  {table}: {count} 条记录")
    
    def run_all(self):
        """运行所有迁移任务"""
        print("\n" + "=" * 60)
        print("Excel 数据迁移工具")
        print("=" * 60)
        print("\n即将开始数据迁移，请确认：")
        print("  1. 已安装 MySQL 并启动服务")
        print("  2. 已执行 schema.sql 创建表结构")
        print("  3. 已备份原始 Excel 文件")
        print("\n是否继续？(y/n): ", end='')
        
        confirm = input().strip().lower()
        if confirm not in ('y', 'yes', '是'):
            print("已取消迁移")
            return
        
        # 执行各项迁移任务
        self.migrate_product_info()
        self.migrate_logistics_mapping()
        self.migrate_price_mapping()
        self.migrate_sales_owner_mapping()
        
        # 验证数据
        self.verify_data()
        
        print("\n" + "=" * 60)
        print("✅ 数据迁移完成！")
        print("=" * 60)
        print("\n下一步：")
        print("  1. 运行配置向导：python report/config/run_config.py")
        print("  2. 测试映射服务：python report/database/mapping_service.py")
        print("  3. 生成报表：python report/runners/master_runner.py")


if __name__ == "__main__":
    migrator = ExcelDataMigrator()
    migrator.run_all()
