import pandas as pd

# Đọc dữ liệu từ file và đặt tên cột
column_names = ['ID', 'Name', 'Phone', 'Address', 'Product']

df = pd.read_excel('data.xlsx', names=column_names)
df['Price'] = 100
df['SL'] = 1
df['KL'] = 3

# Tạo một cột mới để theo dõi số điện thoại trùng lặp
df.sort_values(by='Phone', inplace=True)
df['Phone_Duplicate'] = df['Phone'].duplicated(keep='last')

def highlight_columns(row):
    if row['Phone_Duplicate']:
        return ['background-color: red' if col in ['ID', 'Name', 'Phone', 'Address','Product','Price','SL','KL'] else '' for col in row.index]
    return [''] * len(row.index)

def concatenate_products(group):
    if len(group) > 1:
        group['SL'] = len(group['Product'].unique())
        group['Product'] = ', '.join(group['Product'])
        group['Price'] = group['Price'].sum() - (group['SL'] - 1) * 35000 


    return group

# Gom nhóm các dòng dựa trên số điện thoại và áp dụng hàm concatenate_products cho từng nhóm
df_fix = df.groupby('Phone').apply(concatenate_products)

# Cập nhật các dòng có số điện thoại trùng lặp
df_fix.loc[df_fix['Phone_Duplicate'] == True] = df.loc[df['Phone_Duplicate'] == True].values

# Nếu bạn muốn áp dụng kiểu màu, bạn có thể sử dụng đoạn mã sau
styled_df = df_fix.style.apply(highlight_columns, axis=1)
styled_df.to_excel('result.xlsx', index=False)
