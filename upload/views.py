import base64
import io
import zipfile

from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import csrf_exempt
import pandas as pd
import openpyxl


def upload(request):
    return render(request, "upload.html")


@csrf_exempt
def process(request):
    try:
        # 获取前端传递的 Excel 文件
        excel_file = request.FILES["file"]

        # 使用 Pandas 读取 Excel 文件
        df = pd.read_excel(excel_file)
        data_dic = df.to_dict(orient='records')

        for i in data_dic:
            print(i)





        data1 = [["Data 1", "Data 2", "Data 3"], ["Data 4", "Data 5", "Data 6"]]
        data2 = [["Data 7", "Data 8", "Data 9"], ["Data 10", "Data 11", "Data 12"]]

        # 创建 DataFrame 对象
        df1 = pd.DataFrame(data1, columns=["Column 1", "Column 2", "Column 3"])
        df2 = pd.DataFrame(data2, columns=["Column 1", "Column 2", "Column 3"])

        # 将处理后的数据导出为 Excel 文件
        output = io.BytesIO()
        with zipfile.ZipFile(output, "w") as zf:
            with zf.open("processed_file1.xlsx", "w") as f:
                df1.to_excel(f, index=False, header=True)
            with zf.open("processed_file2.xlsx", "w") as f:
                df2.to_excel(f, index=False, header=True)

        # 创建一个 HttpResponse 对象，设置 content_type 为 "application/zip"
        response = HttpResponse(content_type="application/zip")

        # 设置响应的文件名
        response["Content-Disposition"] = 'attachment; filename="processed_files.zip"'

        # 将打包后的 ZIP 文件保存到 HttpResponse 对象中
        response.write(output.getvalue())

        file_data = base64.b64encode(response.getvalue()).decode()

        return JsonResponse({"success": True, "message": "处理成功", "file": file_data})

    except Exception as e:

        return JsonResponse({"success": False, "message": str(e)})
