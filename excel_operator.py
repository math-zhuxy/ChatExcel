from zhipuai import ZhipuAI
import json
from openpyxl import load_workbook
from openpyxl.cell import Cell
import time

from typing import List
class GLM_MODEL:
    def __init__(self):      
        self.CHATGLM_API_KEY = "e202ab113e4f9e25fff7aac718bc0a1e.FYBjMOyOzwDKj9dA"
        self.FILE_PATH = ""
        self.SAVE_FILE_PATH = "./test.xlsx"
        self.chatglm_client = ZhipuAI(api_key=self.CHATGLM_API_KEY)
        self.FUNCTION_METHOD = [
            {
                "type": "function",
                "function": {
                    "name": "write_excel",
                    "description": "将数据存放到Excel文件中指定的工作表中指定的单元格内",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "sheet":{
                                "type": "string",
                                "description": "要存储的工作表名称，例如 'sheet1'。"
                            },
                            "cell_range": {
                                "type": "string",
                                "description": "要存储的单元格，例如 'C11'。"
                            },
                            "data":{
                                "type": "string",
                                "description": "要存储的数据，例如 'hello world'。"
                            }
                        },
                        "required": ["sheet","cell_range","data"],
                    }
                }
            },
            {
                "type": "function",
                "function": {
                    "name": "read_excel",
                    "description": "读取Excel文件中指定的工作表中指定的单元格范围的数据",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "sheet":{
                                "type": "string",
                                "description": "要读取的工作表名称，例如 'Sheet1'。"
                            },
                            "cell_range": {
                                "type": "string",
                                "description": "要读取的单元格范围，例如 'A1:B10'。"
                            }
                        },
                        "required": ["sheet","cell_range"],
                    }
                }
            },
        ]
        self.messages=[
            {
            "role": "system",
            "content": """
                            1.你是一个Excel的工作助手
                            2.你的输出必须是中文。
                            3.涉及对Excel的操作时尽量调用read_excel或write_excel函数
                            4.不要推测用户的输入，如果用户没有说明你必须问他
                        """
            }
        ]

    def read_excel(self, sheet: str, cell_range: str) -> List[List[str]]:
        print("function read excel called")
        try:
            ws = load_workbook(filename=self.FILE_PATH, data_only=True)[sheet]  # 确保我们得到的是值而不是公式
        except FileExistsError:
            print("文件名不存在")
            return [[]]
        if ':' not in cell_range:
            cell_range=f"{cell_range}:{cell_range}"
        # 读取指定范围的数据
        data = ws[cell_range]

        if isinstance(data,Cell):
            data = [[data]]

        # 将数据转换为列表的列表，并尝试将所有元素转换为整数；如果为空则设为0
        result = []
        for row in data:
            int_row = []
            for cell in row:
                cell_value = cell.value
                try:
                    # 检查单元格是否有内容，如果没有则设为''
                    if cell_value is None or str(cell_value).strip() == '':
                        str_cell = ''
                    else:
                        str_cell = str(cell_value)
                except (ValueError, TypeError) as e:
                    print(f"Error converting cell value '{cell_value}' to integer: {e}")
                    str_cell = ''  # 转换失败也设为空字符串
                int_row.append(str_cell)
            result.append(int_row)
        return result
    
    def is_number(self, num: str) -> bool:
        try:
            float(num)
            return True
        except ValueError:
            return False
    
    def write_excel(self, sheet: str, cell_range: str, data: str) -> str:
        print("function write excel called")
        try:
            wb = load_workbook(filename=self.FILE_PATH, data_only=True)
            if self.is_number(data):
                wb[sheet][cell_range] = float(data)
            else:
                wb[sheet][cell_range] = data
            wb.save(filename=self.SAVE_FILE_PATH)
        except Exception:
            print(f"发生错误{Exception}")
            return "存储数据失败"

        return "存储数据成功"

    def test_readexcel(self, test_sheet: str, test_range: str) -> None:
        print(self.read_excel(test_sheet, test_range))

glm_model = GLM_MODEL()

def excel_operate(user_input: str, file_path: str, progress_callback) -> str:
    glm_model.FILE_PATH = file_path

    glm_model.messages.append({
        "role": "user",
        "content": user_input
    })

    # 创建完请求
    for i in range(11):
        progress_callback(i)
        time.sleep(0.005)

    response = glm_model.chatglm_client.chat.completions.create(
        model = "glm-4-flash", 
        messages = glm_model.messages,
        tools = glm_model.FUNCTION_METHOD,
        tool_choice = "auto"
    )

    glm_model.messages.append(response.choices[0].message.model_dump())

    # 获取到函数调用信息
    print(response)
    for i in range(35):
        progress_callback(i+11)
        time.sleep(0.005)

    if response.choices[0].message.tool_calls:
        tool_call = response.choices[0].message.tool_calls[0]
        args = tool_call.function.arguments
        function_result = {}
        try:
            if tool_call.function.name == "read_excel":
                function_result = glm_model.read_excel(**json.loads(args))
            if tool_call.function.name == "write_excel":
                function_result = glm_model.write_excel(**json.loads(args))

        except Exception:
            print(f"发生错误{Exception}")
            glm_model.messages = []
            return "函数调用错误，请联系开发者或者更改输入"
        
        # 调用完函数
        for i in range(20):
            progress_callback(i+46)
            time.sleep(0.01)
        glm_model.messages.append({
            "role": "tool",
            "content": f"{json.dumps(function_result)}",
            "tool_call_id":tool_call.id
        })

        response = glm_model.chatglm_client.chat.completions.create(
            model = "glm-4-flash",  
            messages = glm_model.messages,
            tools = glm_model.FUNCTION_METHOD,
            tool_choice = "auto"
        )

        glm_model.messages.append(response.choices[0].message.model_dump())

        print(response)

        # 获取到最终请求
        for i in range(35):
            progress_callback(i+66)
            time.sleep(0.005)
    else:
        # 无需调用函数
        for i in range(90):
            progress_callback(i+11)

    return response.choices[0].message.content