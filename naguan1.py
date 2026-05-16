import pandas as pd
import os

# ================= 基础配置 =================
base_file = "base.xlsx"      
platform_file = "plat.csv"
output_file = "work1.xlsx"
unknown_assets_file = "work2.xlsx" # 新增：账外资产输出文件

col_base_ip = "IP地址"
col_deploy_flag = "是否应部署"
col_plat_ip = "主机 IP"
col_plat_status = "主机状态"

status_online = "online"
status_offline = "offline"

def main():
    print(f"1. 正在读取安全平台数据: {platform_file} ...")
    try:
        # 如果报错请尝试 encoding='gbk'
        df_platform = pd.read_csv(platform_file, encoding='utf-8') 
        df_platform.columns = [c.strip() for c in df_platform.columns]
        df_platform[col_plat_ip] = df_platform[col_plat_ip].astype(str).str.strip()
        df_platform[col_plat_status] = df_platform[col_plat_status].astype(str).str.strip().str.lower()
    except Exception as e:
        print(f"读取 CSV 出错: {e}")
        return

    print(f"2. 正在读取基地资产表: {base_file} ...")
    try:
        dict_bases = pd.read_excel(
            base_file, 
            sheet_name=None, 
            engine='openpyxl', 
            engine_kwargs={'data_only': True, 'read_only': True} 
        )
    except Exception as e:
        print(f"读取 Excel 出错: {e}")
        return

    # 用于存放所有已知的基地 IP，进行去重和汇总
    all_known_ips = set()
    
    print("3. 开始对比并生成异常报告...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, df_base in dict_bases.items():
            if df_base.empty: continue
            if col_base_ip not in df_base.columns: continue

            # 记录所有已知 IP (不论是否应部署，只要表里有就算已知)
            ips_in_sheet = df_base[col_base_ip].astype(str).str.strip().unique()
            all_known_ips.update(ips_in_sheet)

            # 过滤“应部署”的资产进行异常分析
            df_filtered = df_base[df_base[col_deploy_flag].astype(str).str.strip() == "是"].copy()
            if df_filtered.empty: continue

            df_filtered[col_base_ip] = df_filtered[col_base_ip].astype(str).str.strip()
            
            merged_df = pd.merge(
                df_filtered,
                df_platform[[col_plat_ip, col_plat_status]],
                left_on=col_base_ip,
                right_on=col_plat_ip,
                how='left'
            )

            mask_uninstalled = merged_df[col_plat_status].isna()
            mask_offline = merged_df[col_plat_status] == status_offline
            result_df = merged_df[mask_uninstalled | mask_offline].copy()

            if not result_df.empty:
                def get_status_label(row):
                    if pd.isna(row[col_plat_status]): return "未安装Agent"
                    return "Agent离线"

                result_df.insert(0, "异常类型", result_df.apply(get_status_label, axis=1))
                if col_plat_ip in result_df.columns: del result_df[col_plat_ip]
                result_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"   -> 异常资产已保存至: {output_file}")

    print("4. 正在分析账外新增资产 (平台有但基地表无)...")
    # 筛选条件：平台是在线状态，且其IP不在已知集合中
    df_unknown = df_platform[
        (df_platform[col_plat_status] == status_online) & 
        (~df_platform[col_plat_ip].isin(all_known_ips))
    ].copy()

    if not df_unknown.empty:
        # 这里尝试根据平台表的某个列（如“所属分组”或“基地名”）来分Sheet输出
        # 如果你的平台表里有区分基地的列，请修改下面的 '所属部门' 为实际列名
        # 如果没有，我们就统一输出到一个 Sheet
        group_col = '所属部门' # 假设平台表有这一列，如果没有请改为 None
        
        with pd.ExcelWriter(unknown_assets_file, engine='openpyxl') as writer_new:
            if group_col in df_unknown.columns:
                for group_name, group_data in df_unknown.groupby(group_col):
                    # Excel Sheet名不能超过31个字符，特殊字符需处理
                    safe_name = str(group_name)[:30]
                    group_data.to_excel(writer_new, sheet_name=safe_name, index=False)
                print(f"   -> 账外资产已按【{group_col}】分类保存至: {unknown_assets_file}")
            else:
                df_unknown.to_excel(writer_new, sheet_name="账外资产清单", index=False)
                print(f"   -> 账外资产已汇总保存至: {unknown_assets_file}")
    else:
        print("   -> 未发现账外新增资产。")

if __name__ == "__main__":
    main()