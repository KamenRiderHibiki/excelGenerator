class DebugInfo:
    """
    输出测试信息
    """
    def __init__(self):
        self.val = 0

    def start(self):
        """
        输出开始信息
        """
        print('测试开始')

    def stop(self):
        """
        输出结束信息
        """
        print('测试结束')

    def print(self, info: str = None) -> None:
        """
        输出测试编号和测试信息
        """
        string = '-'*5 + str(self.val) + '-'*5
        print(string)
        if not not info:  # not x 用于判断info不为0, None, False, 空字符串"", 空列表[], 空字典{}, 空元组()
            print(info)
        self.val += 1
