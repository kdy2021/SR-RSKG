import xlwt
import xlrd
from py2neo import Graph, Node, Relationship,NodeMatcher

#连接neo4j数据库，输入地址、用户名、密码
test_graph = Graph(
    "http://localhost:7474",
    username="neo4j",
    password="123456"
)
test_graph.delete_all()

book = xlrd.open_workbook('rskg.xlsx')
workSheetName = book.sheet_names() #
print("Excel文件包含的表单有："+str(workSheetName))
bridgeStructure = book.sheet_by_name(workSheetName[2])
bridgeStructure2 = book.sheet_by_name(workSheetName[0])
nrows = bridgeStructure.nrows
nrows2 = bridgeStructure2.nrows


#selector = NodeMatcher(test_graph)
def CreatNode(triple,graph):
    a = Node('Class',name= triple[1])
    graph.create(a)
def CreatSubNode(triple,graph):
    a = Node('SubClass',name= triple[1])
    graph.create(a)

def subclassRelationship(triple,graph):
    classname = ["Class","SubClass"]
    list1 = triple[1]
    list2 = triple[2]
    relationship = triple[3]
    className1 = classname[int(triple[4])]
    className2 = classname[int(triple[5])]
    # 利用Python执行CQL语句
    graph.run("match(a:%s),(b:%s)  where a.name='%s' and b.name='%s'  CREATE (a)-[r:%s]->(b)"%(className1,className2,list1,list2,relationship))

#
# for row in range(70):
#     node = bridgeStructure2.row_values(row)
#     print(node)
#     CreatNode(node,test_graph)
# for row in range(70,nrows2):
#     node = bridgeStructure2.row_values(row)
#     print(node)
#     CreatSubNode(node,test_graph)
for row in range(nrows2):
    node = bridgeStructure2.row_values(row)
    if node[2] == 0:
        CreatNode(node, test_graph)
    else:
        CreatSubNode(node, test_graph)


for row in range(nrows):
    a = bridgeStructure.row_values(row)
    subclassRelationship(a,test_graph)
    print(a,test_graph)
    print('创建%d个关系成功'%(row+1))
print("程序运行完成")

#
# print("程序运行完成")
