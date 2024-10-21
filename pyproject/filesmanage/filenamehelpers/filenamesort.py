import re
import os
import sys

class OpFileName:

    @staticmethod
    def simplify_versions(versions):
        # 将输入的字符串按照逗号分隔，生成版本列表，并去除多余的空格
        version_list = [v.strip() for v in versions.split('、')]
        
        # 初始化简化版本列表
        simplified = []
        
        # 初始化一个列表，用于存储连续的版本号
        current_range = []
        
        # 遍历版本列表
        for i in range(len(version_list)):
            if not current_range:  # 如果current_range为空，添加当前版本
                current_range.append(version_list[i])
            else:
                # 检查当前版本与上一个版本是否为连续版本
                current_version_parts = version_list[i].split('.')
                prev_version_parts = current_range[-1].split('.')
                
                # 比较最后一部分是否相差1，且前面部分是否相同
                iscontinuous = False

                try:
                    iscontinuous = (len(
                        
                    ) == len(prev_version_parts) and
                        current_version_parts[:-1] == prev_version_parts[:-1] and
                        int(current_version_parts[-1]) == int(prev_version_parts[-1]) + 1)
                except Exception as e:
                    iscontinuous = False
                if iscontinuous:
                    current_range.append(version_list[i])  # 如果连续，加入current_range
                else:
                    # 如果不连续，结束当前range，并存储到simplified
                    if len(current_range) > 2:  # 如果连续超过两个，才简写为范围
                        simplified.append(f"{current_range[0]}~{current_range[-1]}")
                    else:
                        # 如果是两个连续版本或单独版本，用“、”隔开
                        simplified.extend(current_range)
                    current_range = [version_list[i]]  # 开始一个新的range
        
        # 最后一次的range处理
        if current_range:
            if len(current_range) > 2:  # 如果最后一段有超过两个版本连续
                simplified.append(f"{current_range[0]}~{current_range[-1]}")
            else:
                simplified.extend(current_range)  # 两个连续版本或单独版本
        
        return '、'.join(simplified)