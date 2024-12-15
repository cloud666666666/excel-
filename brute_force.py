import msoffcrypto
import io
import pandas as pd
def try_decrypt_with_passwords(file_path, password_book_path):
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as f:
        office_file = msoffcrypto.OfficeFile(f)
        # 打开密码本逐行读取密码
        with open(password_book_path, 'r',encoding='utf-8') as pw_file:
            for password in pw_file:
                password = password.strip()  # 去掉可能的换行或空格
                try:
                    office_file.load_key(password=password)  # 加载密码
                    office_file.decrypt(decrypted)  # 解密文件
                    print(f"成功解密，使用密码: {password}")
                    decrypted.seek(0)  # 重置流位置
                    return decrypted, password  # 返回解密后的流和正确的密码
                except Exception:
                    continue  # 密码不正确，尝试下一个
    raise ValueError("无法解密文件，密码本中没有正确的密码")
def open_encrypted_xls(file_path, password_book_path):
    """读取加密的 .xls 文件"""
    # 尝试解密
    decrypted, correct_password = try_decrypt_with_passwords(file_path, password_book_path)

    # 使用 Pandas 读取解密后的内容
    df = pd.read_excel(decrypted, engine='openpyxl')  # 用 pandas 读取 Excel 数据
    return df, correct_password
# 示例使用
file_path = '1.xls'  # 替换为你的加密文件路径
password_book_path = 'pwdbook.txt'  # 替换为你的密码本路径

try:
    data, used_password = open_encrypted_xls(file_path, password_book_path)
    print(f"解密成功，使用的密码是: {used_password}")
    print(data)
except Exception as e:
    print(f"解密失败: {e}")
