import pandas as pd
import os

# ================= 基础配置 =================
base_file = "base.xlsx"      
platform_file = "plat.csv"
output_file = "work.xlsx"

col_base_ip = "IP地址"
col_deploy_flag = "是否应部署"  # 新增：应部署判断列
col_plat_ip = "主机 IP"
col_plat_status = "主机状态"

status_online = "online"
status_offline = "offline"

def main():
    print(f"1. 正在读取安全平台数据: {platform_file} ...")
    try:
        # 注意：如果CSV依然报错，请将 utf-8 改为 gbk
        df_platform = pd.read_csv(platform_file, encoding='utf-8') 
        df_platform.columns = [c.strip() for c in df_platform.columns]
        df_platform[col_plat_ip] = df_platform[col_plat_ip].astype(str).str.strip()
        df_platform[col_plat_status] = df_platform[col_plat_status].astype(str).str.strip().str.lower()
        df_platform = df_platform[[col_plat_ip, col_plat_status]]
    except Exception as e:
        print(f"读取 CSV 出错: {e}")
        return

    print(f"2. 正在读取基地资产表: {base_file} ...")
    try:
        # 使用 read_only=True 解决之前的 openpyxl 样式报错问题
        dict_bases = pd.read_excel(
            base_file, 
            sheet_name=None, 
            engine='openpyxl', 
            engine_kwargs={'data_only': True, 'read_only': True} 
        )
    except Exception as e:
        print(f"读取 Excel 出错: {e}")
        return

    print("3. 开始对比并生成报告...")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            issues_found_total = 0
            
            for sheet_name, df_base in dict_bases.items():
                if df_base.empty: continue
                
                # 检查必要的列是否存在
                if col_base_ip not in df_base.columns or col_deploy_flag not in df_base.columns:
                    print(f"  [跳过] Sheet '{sheet_name}' 缺少关键列（'{col_base_ip}' 或 '{col_deploy_flag}'）")
                    continue

                # --- 核心修改：前置过滤 ---
                # 只保留“是否应部署”列为“是”的行。同时处理掉空格和大小写差异
                df_filtered = df_base[df_base[col_deploy_flag].astype(str).str.strip() == "是"].copy()
                
                if df_filtered.empty:
                    print(f"  -> {sheet_name}: 无需部署的资产或数据为空")
                    continue

                # 数据清洗：IP去空格
                df_filtered[col_base_ip] = df_filtered[col_base_ip].astype(str).str.strip()

                # 合并表格 (以过滤后的基地表为主)
                merged_df = pd.merge(
                    df_filtered,
                    df_platform,
                    left_on=col_base_ip,
                    right_on=col_plat_ip,
                    how='left'
                )

                # 判断逻辑：在“应部署”的前提下，找不到状态(NaN)或状态为offline
                mask_uninstalled = merged_df[col_plat_status].isna()
                mask_offline = merged_df[col_plat_status] == status_offline
                
                result_df = merged_df[mask_uninstalled | mask_offline].copy()

                if not result_df.empty:
                    def get_status_label(row):
                        status = row[col_plat_status]
                        if pd.isna(status): return "未安装Agent"
                        elif status == status_offline: return "Agent离线"
                        return "其他异常"

                    result_df.insert(0, "异常类型", result_df.apply(get_status_label, axis=1))
                    
                    # 移除冗余的平台IP列
                    if col_plat_ip in result_df.columns: 
                        del result_df[col_plat_ip]

                    result_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    count = len(result_df)
                    issues_found_total += count
                    print(f"  -> {sheet_name}: 发现 {count} 个异常资产")
                else:
                    print(f"  -> {sheet_name}: 需部署资产状态均正常")

        print("-" * 30)
        if issues_found_total > 0:
            print(f"处理完成！异常清单已保存至: {output_file}")
        else:
            print("完美！所有应部署的资产均已在线。")
            
    except Exception as e:
        print(f"处理过程中发生错误: {e}")

if __name__ == "__main__":
    main()