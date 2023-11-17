import base64
import io
import zipfile
import re
import ipaddress

from upload.constants import *
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pandas as pd
import openpyxl

def is_valid_ipv4(address):    #判断是否为ipv4地址
    # IPv4正则表达式
    pattern = r'^((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])$'
    if re.match(pattern, address):
        return True
    else:
        return False

def is_valid_ipv6(address):     #判断是否为ipv6地址
    try:
        ipaddress.IPv6Address(address)
        return True
    except ipaddress.AddressValueError:
        return False

def is_valid_CIDR(address):    #判断是否为CIDR地址
    # CIDR正则表达式
    pattern = r'^((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])/(?:[0-9]|[1-2][0-9]|3[0-2])$'
    if re.match(pattern, address):
        return True
    else:
        return False

def get_cidr_list(cidr):    #根据CIDR地址生成对应的ip列表
    ip_list = []
    network = ipaddress.ip_network(cidr,strict=False)
    for ip in network:
        ip_list.append(str(ip))
    return ip_list

def is_ipv4_range(ip_range):    #判断是否为ipv4网段，例：192.168.1.1-192.168.1.10
    # ipv4网段正则表达式
    pattern = r'^((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])-((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])$'
    if re.match(pattern, ip_range):
        return True
    else:
        return False

def get_ipv4_range_list(ip_range):     #生成一个ipv4网段中的所有ip，返回包含所有ip的列表
    start_ip, end_ip = ip_range.split('-')
    start_ip_parts = list(map(int, start_ip.split('.')))
    end_ip_parts = list(map(int, end_ip.split('.')))

    ip_range_list = []
    while start_ip_parts <= end_ip_parts:
        ip = '.'.join(map(str, start_ip_parts))
        ip_range_list.append(ip)

        start_ip_parts[3] += 1
        for i in range(3, 0, -1):
            if start_ip_parts[i] == 256:
                start_ip_parts[i] = 0
                start_ip_parts[i - 1] += 1

    return ip_range_list

def is_ipv4_range_2(ip_range) :    #判断是否为第2种ipv4网段，例：192.168.1.1-10
    # 第2种ipv4网段正则表达式
    pattern = r'^((25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])-(25[0-5]|2[0-4][0-9]|1[0-9][0-9]|[1-9][0-9]|[0-9])$'
    if re.match(pattern, ip_range):
        return True
    else:
        return False

def get_ipv4_range_2_list(ip_range):       #生成一个第2种ipv4网段中的所有ip，返回包含所有ip的列表
    start_ip, end_ip = ip_range.split('-')
    start_ip_parts = list(map(int, start_ip.split('.')))
    end_ip = int(end_ip)

    ip_range_list = []
    while start_ip_parts[3] <= end_ip:
        ip = '.'.join(map(str, start_ip_parts))
        ip_range_list.append(ip)
        start_ip_parts[3] += 1

    return ip_range_list


def upload(request):
    return render(request, "upload.html")


@csrf_exempt
def process(request):
    try:
        # 获取前端传递的 Excel 文件
        excel_file = request.FILES["file"]

        # 使用 Pandas 读取 Excel 文件
        df = pd.read_excel(excel_file,dtype=str)

        df.fillna('', inplace=True)

        file_name=excel_file.name[0:-5]

        data_dic = df.to_dict(orient='records')

        # for i in data_dic:
        #     print(i)

        # 获取映射关系
        df2 = pd.read_excel(r'映射.xlsx', dtype=str, header=None)
        data_dic2 = df2.values.tolist();
        # print(data_dic2)
        mapper = dict()
        for i in data_dic2:
            mapper["0" + i[1]] = i[0]
        # print(mapper)

        # 待插入表的列名，按顺序存储的
        headlist = [ASSET_NAME, ASSET_OBJECT_NAME, ASSET_PLATFORM_ID, ASSET_LOCAL_ID, IP_ADDRESS, DOMAIN_IT_BELONGS,
                    WHETHER_TO_TROUBLESHOOT, REASON_FOR_NOT_TROUBLESHOOTING, INTERNAL_AND_EXTERNAL_NETWORK_ASSETS,
                    PORT_PROTOCOL, NETWORK_UNIT, AGGREGATED_ASSET_QUANTITY, OPERATING_SYSTEM, MIDDLEWARE,
                    APPLICATION, HARDWARE, CONTACT_INFORMATION, DOMAIN_NAME, PORT, SERVICE, DATABASE, CPU, MEMORY,
                    SWITCH_CHIP, PROVINCE, OPERATOR, START_TIME, END_TIME]

        mapper_columns = [OPERATING_SYSTEM,
                          MIDDLEWARE,
                          APPLICATION,
                          HARDWARE,
                          SERVICE,
                          DATABASE]

        ip_error_header = [ASSET_NAME,
                           ASSET_OBJECT_NAME,
                           ASSET_PLATFORM_ID,
                           IP_ADDRESS,
                           INTERNAL_AND_EXTERNAL_NETWORK_ASSETS,
                           PORT_PROTOCOL,
                           NETWORK_UNIT,
                           CONTACT_INFORMATION,
                           DOMAIN_NAME,
                           PORT,
                           CPU,
                           MEMORY,
                           SWITCH_CHIP,
                           PROVINCE,
                           OPERATOR,
                           START_TIME,
                           END_TIME,
                           CLEANED_NAME_MANUFACTURER_VERSION,
                           ASSET_TYPE]

        ip_error_list = []

        res = dict()  # 存放处理结果，格式为：元组对应列表，元组为聚合条件，列表为用于比较的列的最终结果以及出现次数

        for dic in data_dic:  # 遍历所有行
            exclude_keys = [IP_ADDRESS, PORT]
            values_dic = {key: value for key, value in dic.items() if
                          key not in exclude_keys}  # 不包含ip和port列的字典部分，用于后续拼接
            # print(values_dic)

            ip = dic[IP_ADDRESS]
            # ip_list = list(set(substr.strip() for substr in re.split(r'[,:;|]', ip) if substr.strip()))   #拆开ip后的列表

            ip_list = []
            flag_jump = False
            for substr in re.split(r'[，,;|]', ip):
                sstr = substr.strip()
                if sstr != "":
                    if is_valid_ipv4(sstr) or is_valid_ipv6(sstr):  # 合法IPV4或IPV6地址
                        ip_list.append(sstr)
                    elif is_valid_CIDR(sstr):  # 合法CIDR地址
                        # print("-----------------CIDR-----------------")
                        ip_list.extend(get_cidr_list(sstr))
                    elif is_ipv4_range(sstr):  # 合法IPV4网段
                        tmp = get_ipv4_range_list(sstr)
                        if len(tmp) == 0:  # IPV4网段中前一个IP严格大于后一个IP，也认为是乱码
                            # print("-----------------IP乱码-----------------")
                            flag_jump = True
                        else:
                            # print("-----------------IPV4网段-----------------")
                            ip_list.extend(tmp)
                    elif is_ipv4_range_2(sstr):  # 合法第2种ipv4网段
                        tmp = get_ipv4_range_2_list(sstr)
                        if len(tmp) == 0:  # 第2种IPV4网段中前一个IP的最后一位严格大于后一个数字，也认为是乱码
                            # print("-----------------IP乱码-----------------")
                            flag_jump = True
                        else:
                            # print("-----------------第2种IPV4网段-----------------")
                            ip_list.extend(tmp)
                    else:
                        # print("-----------------IP乱码-----------------")
                        flag_jump = True
            if flag_jump:
                tmp = []
                for i in ip_error_header:
                    tmp.append(dic[i])
                ip_error_list.append(tmp)
                continue

            ip_list = list(set(ip_list))

            # print(ip_list)

            port = dic[PORT]
            if port == "":
                filtered_port_list = [""]
            else:
                port_list = list(set(substr.strip() for substr in re.split(r'[，,:;|]', port) if substr.strip()))
                filtered_port_list = [substr for substr in port_list if
                                      int(substr) >= 1 and int(substr) <= 65535]  # 拆开port后的列表
            # print(filtered_port_list)

            # 把目标表中不存在的列放进字典，下面不插入聚合值那一列
            values_dic[ASSET_LOCAL_ID] = values_dic[ASSET_PLATFORM_ID]
            values_dic[DOMAIN_IT_BELONGS] = ""
            values_dic[WHETHER_TO_TROUBLESHOOT] = ""
            values_dic[REASON_FOR_NOT_TROUBLESHOOTING] = ""
            for i in mapper_columns:
                if mapper[values_dic[ASSET_TYPE]] == i:
                    values_dic[i] = values_dic[CLEANED_NAME_MANUFACTURER_VERSION]
                else:
                    values_dic[i] = ""

            selected_columns = [IP_ADDRESS,
                                PORT,
                                PORT_PROTOCOL,
                                OPERATING_SYSTEM,
                                MIDDLEWARE,
                                APPLICATION,
                                HARDWARE,
                                SERVICE,
                                DATABASE]  # 按照这些列去聚合
            for x in ip_list:
                for y in filtered_port_list:
                    d = {IP_ADDRESS: x, PORT: y}
                    d.update(values_dic)  # values_dic代表了原先的一行拆开后的每一行
                    result = tuple(d[column] for column in selected_columns)
                    # print(result)
                    if result not in res:
                        res[result] = [d, 1]
                    else:
                        res[result] = [d if res[result][0][START_TIME] < d[START_TIME] else res[result][0],
                                       res[result][1] + 1]

        process_result=[]
        for v in res.values():
            v[0][AGGREGATED_ASSET_QUANTITY] = v[1]
            lis = []
            for i in headlist:
                lis.append(v[0][i])

            process_result.append(lis)


        # 创建 DataFrame 对象
        df1 = pd.DataFrame(process_result, columns=headlist)
        df2 = pd.DataFrame(ip_error_list, columns=ip_error_header)

        # 将处理后的数据导出为 Excel 文件
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as zf:
            with zf.open(f"{file_name}_结果.xlsx", "w") as f:
                df1.to_excel(f, index=False, header=True)
            with zf.open(f"{file_name}_ip乱码.xlsx", "w") as f:
                df2.to_excel(f, index=False, header=True)

        # 创建一个 HttpResponse 对象，设置 content_type 为 "application/zip"
        response = HttpResponse(content_type="application/zip")

        # 设置响应的文件名
        response["Content-Disposition"] = f'attachment; filename="{file_name}_处理结果.zip"'

        # 将打包后的 ZIP 文件保存到 HttpResponse 对象中
        response.write(output.getvalue())

        file_data = base64.b64encode(response.getvalue()).decode()

        return JsonResponse({"success": True, "message": "处理成功", "file": file_data})

    except KeyError as keye:

        return JsonResponse({"success": False, "message": "该列不存在："+str(keye).strip("'")+"，请检查所给文件的列名或没有选择上传文件"})

    except Exception as e:

        return JsonResponse({"success": False, "message": "出错："+str(e)})
