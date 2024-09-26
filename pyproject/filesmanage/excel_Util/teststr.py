def simplify_versions(versions):
    version_list = versions.split('、')  # 按 '、' 分割成列表
    simplified = []
    start = None

    for i in range(len(version_list)):
        if start is None:  # 初始化起始版本
            start = version_list[i]
        
        # 如果下一个版本不连续，或者已经是最后一个版本
        if i == len(version_list) - 1 or not is_consecutive(version_list[i], version_list[i+1]):
            if start == version_list[i]:  # 没有连续的版本
                simplified.append(start)
            else:  # 有连续的版本
                simplified.append(f"{start}~{version_list[i]}")
            start = None  # 重置起始版本

    return '、'.join(simplified)

def is_consecutive(v1, v2):
    # 判断版本号 v1 和 v2 是否连续
    v1_parts = list(map(int, v1.split('.')))
    v2_parts = list(map(int, v2.split('.')))

    # 仅处理版本号第三部分
    if v1_parts[:2] == v2_parts[:2] and v2_parts[2] == v1_parts[2] + 1:
        return True
    return False

# 示例
versions = "1.0.1、1.0.2、1.0.3、1.0.4、1.0.8"
result = simplify_versions(versions)
print(result)  # 输出: 1.0.1~1.0.4、1.0.8
