import pandas as pd
import numpy as np
import re

def clean_amount(val):
    if pd.isna(val) or val == 'nan':
        return 0.0
    s = str(val)
    cleaned = re.sub(r'[^\d\.-]', '', s)
    if cleaned:
        try:
            return float(cleaned)
        except:
            return 0.0
    return 0.0

def process_excel(input_path, output_path):
    print(f"Reading {input_path} with dtype=str...")
    df = pd.read_excel(input_path, sheet_name=0, dtype=str)
    original_df = df.copy()
    
    df['纯金额'] = df['变动金额'].apply(clean_amount)
    
    processed_rows = []
    
    grouped = df.groupby('订单号')
    
    for order_id, group in grouped:
        main_fee_rows = group[group['费用项'] == '正向配送费']
        
        # Rule 1: Must have 正向配送费, otherwise delete
        if len(main_fee_rows) == 0:
            continue
            
        total_amount = group['纯金额'].sum()
        
        # Rule 2: Check if split is needed
        if len(main_fee_rows) > 1 and abs(total_amount) > 190:
            # SPLIT ORDER
            other_fee_rows = group[group['费用项'] != '正向配送费']
            
            main_amounts = main_fee_rows['纯金额'].tolist()
            total_main_fee = sum(main_amounts)
            total_other_fee = other_fee_rows['纯金额'].sum()
            
            other_items = list(dict.fromkeys(other_fee_rows['费用项'].dropna().astype(str)))
            other_items_str = ", ".join(other_items) if other_items else "无其他费用"
            
            num_splits = len(main_amounts)
            allocated_others = []
            
            # Calculate allocated other fees with exact penny-rounding balance on the last sub-order
            current_allocated_sum = 0.0
            for i in range(num_splits):
                if i < num_splits - 1:
                    ratio = main_amounts[i] / total_main_fee if total_main_fee != 0 else (1.0 / num_splits)
                    alloc = round(total_other_fee * ratio, 2)
                    current_allocated_sum += alloc
                else:
                    alloc = round(total_other_fee - current_allocated_sum, 2)
                allocated_others.append(alloc)
                
            for i in range(num_splits):
                sub_idx = i + 1
                first_row = main_fee_rows.iloc[i].copy()
                
                main_amt = main_amounts[i]
                other_alloc = allocated_others[i]
                sub_total = round(main_amt + other_alloc, 2)
                
                ratio_pct = (main_amt / total_main_fee * 100) if total_main_fee != 0 else (100.0 / num_splits)
                
                first_row['订单号'] = f"{order_id}_{sub_idx}"
                first_row['变动金额'] = sub_total
                first_row['费用项'] = f"总金额(正向配送费, {other_items_str}[按比例分摊])"
                
                detail_str = (f"正向配送费: {main_amt:.2f}; "
                              f"其他费用总计({total_other_fee:.2f})按比例({ratio_pct:.2f}%)分配: {other_alloc:.2f}; "
                              f"小计: {sub_total:.2f}")
                first_row['金额计算明细'] = detail_str
                
                processed_rows.append(first_row)
        else:
            # UNSPLIT ORDER
            first_row = group.iloc[0].copy()
            first_row['变动金额'] = round(total_amount, 2)
            
            items = list(dict.fromkeys(group['费用项'].dropna().astype(str)))
            first_row['费用项'] = f"总金额({', '.join(items)})"
            
            # Build detail string for unsplit order
            fee_details = []
            for item, item_group in group.groupby('费用项'):
                item_sum = item_group['纯金额'].sum()
                fee_details.append(f"{item}: {item_sum:.2f}")
            
            detail_str = f"{'; '.join(fee_details)}; 合计: {total_amount:.2f}"
            first_row['金额计算明细'] = detail_str
            
            processed_rows.append(first_row)
            
    result_df = pd.DataFrame(processed_rows)
    if '纯金额' in result_df.columns:
        result_df = result_df.drop(columns=['纯金额'])
        
    print(f"Writing to {output_path}...")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        original_df.to_excel(writer, sheet_name='Sheet1', index=False)
        result_df.to_excel(writer, sheet_name='Sheet2', index=False)
    print("Done!")

if __name__ == '__main__':
    input_file = '/Users/a171325./Downloads/收支流水20260712230139 (1) 2.xlsx'
    output_file = '/Users/a171325./Downloads/收支流水_processed.xlsx'
    process_excel(input_file, output_file)
