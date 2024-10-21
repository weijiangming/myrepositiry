import re

def find_missing_versions(versions):
    split_versions = [re.split(r'([A-Z]?)$', version) for version in versions]
    groups = {}
    
    for version_parts in split_versions:
        base = version_parts[0]
        letter = version_parts[1]
        if base not in groups:
            groups[base] = []
        groups[base].append(letter)
    
    missing_versions = []
    
    for base, letters in groups.items():
        if letters[0] == '':
            # Find missing numeric versions
            last_part_numbers = sorted(int(base.split('.')[-1]) for base in versions)
            for i in range(last_part_numbers[0], last_part_numbers[-1]):
                if i not in last_part_numbers:
                    #missing_versions.append(f"{base[:-1]}{i}")
                    missing_versions.append(f"{'.'.join(base.split('.')[:-1])}.{i}")
        else:
            # Find missing alphabetic versions
            letters = sorted(letters)
            for i in range(ord(letters[0]), ord(letters[-1])):
                if chr(i) not in letters:
                    missing_versions.append(f"{base}{chr(i)}")

    return list(set(missing_versions))  # Avoid redundant outputs

def find_missing_from_string(version_string):
    # Split the string into individual versions by the separator '、'
    versions = version_string.split('、')

    # Process the list of versions to find missing versions
    return find_missing_versions(versions)

# Example usage
test_strings = [
    "1.0.1、1.0.2、1.0.3、1.0.4、1.0.6",
    "1.0.11、1.0.12、1.0.13、1.0.14、1.0.16",
    "2.0.1A、2.0.1B、2.0.1E",
    "6.3.2A、6.3.2C、6.3.2D",
    "b.0.1、b.0.2、b.0.3、b.0.4、b.0.5、b.0.7",
    "C.0.1、C.0.2、C.0.4"
]
# test_strings = [
#     "1.0.11、1.0.12、1.0.13、1.0.14、1.0.16"
#]

# Process each string and print the missing versions
missing_results = [find_missing_from_string(s) for s in test_strings]
print(missing_results)






# def simplify_versions(versions):
#     version_list = [v.strip() for v in versions.split('、')]

#     simplified = []
#     current_range = []

#     for i in range(len(version_list)):
#         if not current_range:
#             current_range.append(version_list[i])
#         else:
#             current_version_parts = version_list[i].split('.')
#             prev_version_parts = current_range[-1].split('.')

#             # 比较数字部分和字母部分是否连续
#             if (len(current_version_parts) == len(prev_version_parts) and
#                     all(a == b for a, b in zip(current_version_parts[:-1], prev_version_parts[:-1])) and
#                     int(current_version_parts[-2]) == int(prev_version_parts[-2]) and
#                     ord(current_version_parts[-1]) == ord(prev_version_parts[-1]) + 1):
#                 current_range.append(version_list[i])
#             else:
#                 if len(current_range) > 2:
#                     simplified.append(f"{current_range[0]}~{current_range[-1]}")
#                 else:
#                     simplified.extend(current_range)
#                 current_range = [version_list[i]]

#     if current_range:
#         if len(current_range) > 2:
#             simplified.append(f"{current_range[0]}~{current_range[-1]}")
#         else:
#             simplified.extend(current_range)

#     return '、'.join(simplified)

# # 示例测试
# versions = "1.0.1、1.0.2、1.0.3、1.0.4、1.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: 1.0.1~1.0.4、1.0.8

# versions = "2.0.1A、2.0.1B、2.0.1C"
# result = simplify_versions(versions)
# print(result)  # 输出: 1.0.1~1.0.4、1.0.8

# versions = "6.3.2A、6.3.2B、6.3.2C、6.3.2D"
# result = simplify_versions(versions)
# print(result)  # 输出: 1.0.1~1.0.4、1.0.8


# versions = "b.0.1、b.0.2、b.0.3、b.0.4、b.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: b.0.1~b.0.4、b.0.8

# versions = "C.0.1、C.0.2、C.0.4"
# result = simplify_versions(versions)
# print(result)  # 输出: C.0.1、C.0.2、C.0.4

# def simplify_versions(versions):
#     # 将输入的字符串按照逗号分隔，生成版本列表，并去除多余的空格
#     version_list = [v.strip() for v in versions.split('、')]
    
#     # 初始化简化版本列表
#     simplified = []
    
#     # 初始化一个列表，用于存储连续的版本号
#     current_range = []
    
#     # 遍历版本列表
#     for i in range(len(version_list)):
#         if not current_range:  # 如果current_range为空，添加当前版本
#             current_range.append(version_list[i])
#         else:
#             # 检查当前版本与上一个版本是否为连续版本
#             current_version_parts = version_list[i].split('.')
#             prev_version_parts = current_range[-1].split('.')
            
#             # 比较最后一部分是否相差1，且前面部分是否相同
#             if (len(current_version_parts) == len(prev_version_parts) and
#                 current_version_parts[:-1] == prev_version_parts[:-1] and
#                 int(current_version_parts[-1]) == int(prev_version_parts[-1]) + 1):
#                 current_range.append(version_list[i])  # 如果连续，加入current_range
#             else:
#                 # 如果不连续，结束当前range，并存储到simplified
#                 if len(current_range) > 2:  # 如果连续超过两个，才简写为范围
#                     simplified.append(f"{current_range[0]}~{current_range[-1]}")
#                 else:
#                     # 如果是两个连续版本或单独版本，用“、”隔开
#                     simplified.extend(current_range)
#                 current_range = [version_list[i]]  # 开始一个新的range
    
#     # 最后一次的range处理
#     if current_range:
#         if len(current_range) > 2:  # 如果最后一段有超过两个版本连续
#             simplified.append(f"{current_range[0]}~{current_range[-1]}")
#         else:
#             simplified.extend(current_range)  # 两个连续版本或单独版本
    
#     return '、'.join(simplified)



# def simplify_versions(versions):
#     # 将输入的字符串按照逗号分隔，生成版本列表，并去除多余的空格
#     version_list = [v.strip() for v in versions.split('、')]
    
#     # 初始化简化版本列表
#     simplified = []
    
#     # 初始化一个列表，用于存储连续的版本号
#     current_range = []
    
#     # 遍历版本列表
#     for i in range(len(version_list)):
#         if not current_range:  # 如果current_range为空，添加当前版本
#             current_range.append(version_list[i])
#         else:
#             # 检查当前版本与上一个版本是否为连续版本
#             current_version_parts = version_list[i].split('.')
#             prev_version_parts = current_range[-1].split('.')
            
#             # 比较最后一部分是否相差1，且前面部分是否相同
#             if (len(current_version_parts) == len(prev_version_parts) and
#                 current_version_parts[:-1] == prev_version_parts[:-1] and
#                 int(current_version_parts[-1]) == int(prev_version_parts[-1]) + 1):
#                 current_range.append(version_list[i])  # 如果连续，加入current_range
#             else:
#                 # 如果不连续，结束当前range，并存储到simplified
#                 if len(current_range) > 1:
#                     simplified.append(f"{current_range[0]}~{current_range[-1]}")
#                 else:
#                     simplified.append(current_range[0])
#                 current_range = [version_list[i]]  # 开始一个新的range
    
#     # 最后一次的range处理
#     if current_range:
#         if len(current_range) > 1:
#             simplified.append(f"{current_range[0]}~{current_range[-1]}")
#         else:
#             simplified.append(current_range[0])
    
#     return '、'.join(simplified)


# # 示例测试
# versions = "1.0.1、1.0.2、1.0.3、1.0.4、1.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: 1.0.1~1.0.4、1.0.8

# versions = "b.0.1、b.0.2、b.0.3、b.0.4、b.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: b.0.1~b.0.4、b.0.8


# import re

# def simplify_versions(versions):
#     version_list = versions.split('、')  # 按 '、' 分割成列表

#     if len(version_list) < 3:
#         return versions

#     simplified = []
#     start = None

#     for i in range(len(version_list)):
#         if start is None:  # 初始化起始版本
#             start = version_list[i]
        
#         # 如果下一个版本不连续，或者已经是最后一个版本
#         if i == len(version_list) - 1 or not is_consecutive(version_list[i], version_list[i+1]):
#             if start == version_list[i]:  # 没有连续的版本
#                 simplified.append(start)
#             else:  # 有连续的版本
#                 simplified.append(f"{start}~{version_list[i]}")
#             start = None  # 重置起始版本

#     return '、'.join(simplified)

# def is_consecutive(v1, v2):
#     # 判断版本号 v1 和 v2 是否连续
#     v1_parts = extract_version_numbers(v1)
#     v2_parts = extract_version_numbers(v2)

#     if not v1_parts or not v2_parts:  # 如果无法提取到数字版本，返回 False
#         return False

#     # 比较版本号的所有部分，如果前三部分相同并且第四部分是连续的
#     if v1_parts[:2] == v2_parts[:2] and v2_parts[2] == v1_parts[2] + 1:
#         return True
#     return False

# def extract_version_numbers(version):
#     # 提取版本号中的数字部分，忽略非数字部分
#     match = re.findall(r'(\d+)', version)
#     if match:
#         return list(map(int, match))
#     return None

# # 示例
# versions = "b.0.1、b.0.2、b.0.3、b.0.4、b.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: "b.0.1~b.0.4、b.0.8"


# import re

# def simplify_versions(versions):
#     version_list = versions.split('、')  # 按 '、' 分割成列表

#     if len(version_list) < 3:
#         return versions

#     simplified = []
#     start = None

#     for i in range(len(version_list)):
#         if start is None:  # 初始化起始版本
#             start = version_list[i]
        
#         # 如果下一个版本不连续，或者已经是最后一个版本
#         if i == len(version_list) - 1 or not is_consecutive(version_list[i], version_list[i+1]):
#             if start == version_list[i]:  # 没有连续的版本
#                 simplified.append(start)
#             else:  # 有连续的版本
#                 simplified.append(f"{start}~{version_list[i]}")
#             start = None  # 重置起始版本

#     return '、'.join(simplified)

# def is_consecutive(v1, v2):
#     # 判断版本号 v1 和 v2 是否连续
#     v1_parts = extract_version_numbers(v1)
#     v2_parts = extract_version_numbers(v2)

#     if not v1_parts or not v2_parts:  # 如果无法提取到数字版本，返回 False
#         return False

#     # 仅处理版本号第三部分
#     if v1_parts[:2] == v2_parts[:2] and v2_parts[2] == v1_parts[2] + 1:
#         return True
#     return False

# def extract_version_numbers(version):
#     # 提取版本号中的数字部分，忽略字母部分
#     try:
#         return list(map(int, re.findall(r'\d+', version)))
#     except ValueError:
#         return None

# # 示例
# versions = "b.0.1、b.0.2、b.0.3、b.0.4、b.0.8"
# result = simplify_versions(versions)
# print(result)  # 输出: "b.0.1~b.0.4、b.0.8"



# def simplify_versions(versions):
#     version_list = versions.split('、')  # 按 '、' 分割成列表

#     if len(version_list) < 3:
#         return versions

#     simplified = []
#     start = None

#     for i in range(len(version_list)):
#         if start is None:  # 初始化起始版本
#             start = version_list[i]
        
#         # 如果下一个版本不连续，或者已经是最后一个版本
#         if i == len(version_list) - 1 or not is_consecutive(version_list[i], version_list[i+1]):
#             if start == version_list[i]:  # 没有连续的版本
#                 simplified.append(start)
#             else:  # 有连续的版本
#                 simplified.append(f"{start}~{version_list[i]}")
#             start = None  # 重置起始版本

#     return '、'.join(simplified)

# def is_consecutive(v1, v2):
#     # 判断版本号 v1 和 v2 是否连续
#     if v1 == "A.0.1":
#         print(v1)
#     v1_parts = extract_version_numbers(v1)
#     v2_parts = extract_version_numbers(v2)

#     if not v1_parts or not v2_parts:  # 如果无法提取到数字版本，返回 False
#         return False

#     # 仅处理版本号第三部分
#     if v1_parts[:2] == v2_parts[:2] and v2_parts[2] == v1_parts[2] + 1:
#         return True
#     return False

# def extract_version_numbers(version):
#     # 提取版本号中的数字部分，忽略非数字部分
#     try:
#         return list(map(int, filter(lambda x: x.isdigit(), version.split('.'))))
#     except ValueError:
#         return None

# # 示例
# # versions = "1.0.1、1.0.2、1.0.3、1.0.4、1.0.8"
# # result = simplify_versions(versions)
# # print(result)  # 输出: 1.0.1~1.0.4、1.0.8

# versions = "b.0.1、b.0.2、b.0.3、b.0.4、b.0.8"
# result = simplify_versions(versions)
# print(result) 
