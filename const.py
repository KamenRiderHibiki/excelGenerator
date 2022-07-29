# -*- coding: utf-8 -*-
# python 3.x
import sys


class _const:
    """
    实现常量的功能
    该类定义了一个方法__setattr()__,和一个异常ConstError, ConstError类继承自类TypeError
    """
    class ConstError(TypeError):
        pass

    def __setattr__(self, name, value):
        """
        通过调用类自带的字典__dict__, 判断定义的常量是否包含在字典
        中。如果字典中包含此变量，将抛出异常，否则，给新创建的常量赋值。
        """
        if name in self.__dict__:
            raise self.ConstError("Can't rebind const (%s)" % name)
        self.__dict__[name] = value


sys.modules[__name__] = _const()  # 把const类注册到sys.modules这个全局字典中
