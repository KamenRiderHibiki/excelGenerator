import const
import xml.etree.ElementTree as ET

const.TYPE_FOLDER = 'folder'
const.TYPE_EXCEL = 'excel'
const.STR_INPUT = '__INPUT__'


class MultiWayTree:
    """
    多叉树结构实现
    """
    def __init__(self, name=None, nodeType=None):
        self.name = name
        self.type = nodeType
        self.child = []

    def add_child(self, MultiWayTree):
        """
        增加子结点
        """
        self.child.append(MultiWayTree)

    def printTreeNodes(self):
        """
        输出结点信息
        """
        print('{name: ' + self.name + '; type: ' + self.type + '}')
        for c in self.child:
            c.printTreeNodes()


def buildMultiWayTreefromXml(path: str) -> MultiWayTree:
    """
    读取xml创建多叉树
    """
    tree = ET.ElementTree(file=path)
    nodeTree = buildMultiWayTreefromXmlET(tree)
    return nodeTree


def DFSXmlElement(e: ET.Element) -> MultiWayTree:
    """
    深度遍历ElementTree元素
    """
    result = MultiWayTree(e.attrib['name'], e.attrib['type'])
    for child_of_e in e:
        result.add_child(DFSXmlElement(child_of_e))
    return result


def buildMultiWayTreefromXmlET(et: ET) -> MultiWayTree:
    """
    读取xml创建ElementTree
    """
    root = et.getroot()
    return DFSXmlElement(root)
