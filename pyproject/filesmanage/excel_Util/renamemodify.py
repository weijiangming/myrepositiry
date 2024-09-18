#改名
# "城市公共设施规划规范GB504422008",
# "电磁屏蔽室工程施工及质量验收规范GBT511032015",
# "沉井与气压沉箱施工规范"
# 成：
# 《城市公共设施规划规范》GB 50442-2008_条文说明
# 《电磁屏蔽室工程施工及质量验收规范》GB T51103-2015_条文说明
# 《沉井与气压沉箱施工规范》_条文说明

import re

def format_standard_code_final(code):
    # 正则表达式匹配标准名称、标准代码和年份（允许标准代码中间有空格）
    pattern = r'^(.*?)([A-Z]+(?:\s*[A-Z]+)*)\s*(\d+)(?:(\d{4}))?$'
    match = re.search(pattern, code)

    if match:
        # 提取匹配的部分
        name, code_prefix, code_number, year = match.groups()
        # 去除标准代码中的空格
        code_prefix = code_prefix.replace(' ', '')
        # 处理标准代码和年份格式
        if year:
            formatted_code = f"《{name}》{code_prefix} {code_number}-{year}"
        else:
            formatted_code = f"《{name}》{code_prefix} {code_number}"
    else:
        # 如果没有匹配到标准代码和年份，则只添加书名号
        formatted_code = f"《{code}》"

    return formatted_code

# 测试
# codes = [
#     "城市公共设施规划规范GB504422008",
#     "电磁屏蔽室工程施工及质量验收规范GBT511032015",
#     "沉井与气压沉箱施工规范"
# ]

# formatted_codes = [format_standard_code_final(code) for code in codes]
# formatted_codes
