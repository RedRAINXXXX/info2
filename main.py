import re
import pandas as pd
import os
import glob
import natsort
from tqdm import tqdm
import sys
import os.path

from win32com.client import Dispatch

#initialization
app = Dispatch('Word.Application')
excel = Dispatch('Excel.Application')
data_list = []
label_list = []
color_dict = {'橙色':49407,'蓝色':15773696,'绿色':5287936,'浅绿':5296274}

try:
    if os.path.isfile('conf.xlsx'):
        conf_path = "conf.xlsx"
    else:
        conf_path = input("默认配置文件不存在，请输入配置文件路径：")
    conf = pd.read_excel(conf_path, header=0, index_col=0)
    doc1 = app.Documents.Open(conf.loc['doc_source'][0])
    doc1.ActiveWindow.View.ShowHiddenText = False
    doc1.Activate()
    xbook = excel.Workbooks.Open(conf.loc['data_excel'][0])

    app.visible = True if conf.loc['word_visible'][0] == 1 else False
    excel.visible = True if conf.loc['excel_visible'][0] == 1 else False

    loc_num = conf.loc['loc_num'][0]
    for i in range(loc_num):
        data_row = conf.loc['data_loc_{}'.format(i + 1)]
        data_list.append((data_row[0], (int(data_row[1]), int(data_row[2])), (int(data_row[3]), int(data_row[4]))))
        label_row = conf.loc['label_loc_{}'.format(i + 1)]
        label_list.append((label_row[0], (int(label_row[1]), int(label_row[2])), (int(label_row[3]), int(label_row[4]))))
except Exception as errmsg:
    print(errmsg)

#工作表名称 (起始行，起始列) (结尾行，结尾列)
# data_list = [('替换要素',(2,3),(179,3))]
# label_list = [('替换要素',(2,1),(179,1))]


def fiFindByWildcard(wildcard):
    return natsort.natsorted(glob.glob(wildcard, recursive=True))
def replace_all(oldstr, newstr, regrex = False):
    app.Selection.Find.Execute(
        oldstr, False, False, regrex, False, False, True, 1, False, newstr, 2)

# app.Selection.Find.Execute('\{\$(*)\}', False, False, True, False, False, True, 1, False, '\\1', 2)
#Replace

#1 Remove Color
def remove_color(mode='conf'):
    """
    删除指定的颜色
    :param mode: input: input conf: conf file
    :return:
    """
    colorname = None
    if mode == 'input':
        colorname = input("请输入要删除的颜色，（支持颜色：橙色、蓝色、绿色、浅绿）：")
    elif mode == 'conf':
        colorname = conf.loc['1_colorname'][0]
    app.Selection.Find.Font.Color = color_dict[colorname]
    app.Selection.Find.Execute(FindText='', MatchCase=False, MatchWholeWord=False, MatchWildcards = False,
                               MatchSoundsLike = False, MatchAllWordForms = False, Forward = True,
                               Wrap = 1, Format = True, ReplaceWith='', Replace=2)
    print("Func 1 Done!")

def set_find(ClearFormatting=True, Text="", ReplacementText="",Forward=True,
             Wrap = 1,Format = False,MatchCase = False,MatchWholeWord = False,
             MatchByte = True,MatchAllWordForms = False,MatchSoundsLike = False,MatchWildcards = True):
    if ClearFormatting:
        app.Selection.Find.ClearFormatting()
    app.Selection.Find.Text = Text
    app.Selection.Find.Replacement.Text = ReplacementText
    app.Selection.Find.Forward = Forward
    app.Selection.Find.Wrap = Wrap
    app.Selection.Find.Format = Format
    app.Selection.Find.MatchCase = MatchCase
    app.Selection.Find.MatchWholeWord = MatchWholeWord
    app.Selection.Find.MatchByte = MatchByte
    app.Selection.Find.MatchAllWordForms = MatchAllWordForms
    app.Selection.Find.MatchSoundsLike = MatchSoundsLike
    app.Selection.Find.MatchWildcards = MatchWildcards

def hide_change(hidden = False,name = None):
    set_find(Text="\{T%s_*_T%s\}" % (name, name))
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        app.Selection.Font.Hidden = hidden

def all_com(now,rest):
    if rest != '':
        all_com(now, rest[1:])
        all_com(now+rest[0], rest[1:])
    else:
        if len(now) != 0:
            hide_change(hidden =True, name = ';'.join(now))

#2 Hide Template
def hide_all_T(mode='conf'):
    """
    隐藏所有的数字模板
    :param mode: input: input conf: conf file
    :return:
    """
    t_num = None
    if mode == 'input':
        t_num = int(input("请输入模板总类数："))
    elif mode == 'conf':
        t_num = int(conf.loc['t_num'][0])

    doc1.ActiveWindow.View.ShowHiddenText = True
    tl = [str(i) for i in range(1,t_num + 1)]
    all_com(now='',rest=''.join(tl))
    doc1.ActiveWindow.View.ShowHiddenText = False

    print("Func 2 Done!")

def chosen_copy(now,rest,chosen):
    if rest != '':
        chosen_copy(now, rest[1:], chosen)
        chosen_copy(now+rest[0], rest[1:], chosen)
    else:
        if len(now) != 0 and chosen in now:
            name = ';'.join(now)
            replace_all("\{T%s_(*)_T%s\}" % (name, name),"{P_\\1_P}{T%s_\\1_T%s}" % (name, name),regrex=True)
            # replace_all("^p^p", "^p")

#3 Copy Chosen Template
def copy_chosen_T(mode='conf'):
    """
    将选择的模板复制一份P标记副本
    :param mode: input: input conf: conf file
    :return:
    """
    t_num = None
    chosen = None
    if mode == 'input':
        t_num = int(input("请输入模板总类数："))
        chosen = int(input("请输入要复制的模板："))
    elif mode == 'conf':
        t_num = int(conf.loc['t_num'][0])
        chosen = int(conf.loc['3_chosen_t'][0])

    tl = [str(i) for i in range(1,t_num + 1)]
    chosen_copy(now='', rest=''.join(tl), chosen=str(chosen))

    print("Func 3 Done!")

#4 Restore from P mode
def restore(mode):
    """
    删除所有的P标记模板
    :return:
    """
    doc1.ActiveWindow.View.ShowHiddenText = True
    app.Selection.WholeStory()
    replace_all('\{P_(*)_P\}', '', regrex=True)
    app.Selection.Font.Hidden = False
    doc1.ActiveWindow.View.ShowHiddenText = False

    print("Func 4 Done!")

#5 Replace Single File
def replace_doc1(mode):
    """
    从data_list,label_list中读取数据，替换doc_source
    :return:
    """
    replace_elements(doc1.Name)

    print("Func 5 Done!")

def replace_elements(filename):
    for i,l in zip(data_list,label_list):
        row_num = i[2][0]-i[1][0]+1
        col_num = i[2][1]-i[1][1]+1
        with tqdm(total=row_num*col_num) as pbar:
            pbar.set_description(filename+'_'+i[0])
            for row in range(row_num):
                for col in range(col_num):
                    data = xbook.sheets[i[0]].Cells(i[1][0]+row,i[1][1]+col).text.strip()
                    label = xbook.sheets[l[0]].Cells(l[1][0]+row,l[1][1]+col).text.strip()
                    if data!='' and data!='#DIV/0!' and data!='-':
                        replace_all(label,data)
                    pbar.update(1)

#6 Remove Label
def remove_mark(mode='conf'):
    """
    去除文件的所有指定标记，并保留一份备份文件
    :param mode: input: input conf: conf file
    :return:
    """

    labels = None
    if mode == 'input':
        labels = input("请输入要去除的标记（例：P或P;C;D）：")
    elif mode == 'conf':
        labels = conf.loc['6_labels'][0]

    doc1.Save()
    doc1.SaveAs(os.path.dirname(os.path.abspath(__file__)) + '/{}_remove.docx'.format(doc1.Name.split('.')[0]))

    if re.match('^([PCD];)+[PCD]$', labels):
        for label in labels.split(";"):
            replace_all('\{%s_(*)_%s\}' % (label, label), '\\1', regrex=True)
    elif re.match("^[PCD]$", labels):
        replace_all('\{%s_(*)_%s\}' % (labels, labels), '\\1', regrex=True)

    print("Func 6 Done!")

#7 Remove Specific Row of certain tables
def remove_condition_row(mode='conf'):
    """
    去除特定标记范围内的“空”行
    :param mode: input: input conf: conf file
           enum: ''：去除C标记范围内所有表格的全为空的行 '-'：去除D标记范围内所有表格的从第二列全为'-'的行
    :return:
    """
    # doc1.ActiveWindow.View.ShowHiddenText = True
    # default enum ''

    enum = None
    if mode == 'input':
        enum = input("请输入要去除的行类型（例：'空'或'-'）：")
    elif mode == 'conf':
        enum = conf.loc['7_enum'][0]

    if enum == '空':
        enum = ''
        col_begin = 1
        set_find(Text="\{C_(*)_C\}")
    elif enum == '-':
        col_begin = 2
        set_find(Text="\{D_(*)_D\}")
    else:
        return

    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        for table in app.Selection.Range.tables:
            if table.Cell(1,1).range.font.hidden == -1:
                continue
            row_index = 1
            for i in range(1, table.rows.count + 1):
                flag = True
                for j in range(col_begin, table.columns.count + 1):
                    try:
                        cell_str = table.Cell(row_index,j).range.text.split('\r')[0]
                        if cell_str != enum:
                            flag = False
                            break
                    except BaseException:
                        flag = False
                        break
                if flag:
                    doc1.Range(table.Cell(row_index,1).range.start,table.Cell(row_index,table.columns.count).range.end).cells.Delete(2)
                    row_index -= 1
                row_index += 1

    print("Func 7 Done!")
    # doc1.ActiveWindow.View.ShowHiddenText = False

#8 Remove paragraphs with specific trigger
def delete_trigger_para(mode='conf'):
    """
    去除带有特定触发字符串的段落
    :param mode: input: input conf: conf file
                 trigger_text: 触发器文字
    :return:
    """
    trigger_text = None
    if mode == 'input':
        trigger_text = input("请输入触发器字符串：")
    elif mode == 'conf':
        trigger_text = conf.loc['8_trigger_text'][0]

    set_find(Text=trigger_text, MatchWildcards=False)
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        app.selection.paragraphs[0].Range.Delete()

    print("Func 8 Done!")


def level_1_condition():
    set_find(Text="\{E(*)_")
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        label_text = app.Selection.text[2:-1]

        #CONDITION JUDGE
        try:
            cond = 1 if conf.loc[label_text][0] == 1  else 0
        except Exception:
            cond = 0
        #NEED
        if cond == 1:
            oldStr = "\{E%s_(*)_E%s\}" % (label_text, label_text)
            newStr = "{P_\\1_P}{E%s_\\1_E%s}" % (label_text, label_text)
            app.Selection.Find.Execute(oldStr, False, False, True, False, False, True, 1, False, newStr, 1)
            #HIDE T
            start = app.Selection.range.start
            end = app.Selection.range.end
            content_length = (end - start - 12 - 2*len(label_text))//2
            Tstart = content_length + 2 * len(label_text) + 6
            doc1.Range(end-Tstart,end).Font.Hidden = True
        #DONT NEED
        elif cond == 0:
            app.Selection.Find.Text = "\{E%s_(*)_E%s\}" % (label_text, label_text)
            app.Selection.End = app.Selection.Start
            app.Selection.Find.Execute()
            app.selection.Font.Hidden = True

        app.Selection.Start = app.Selection.End
        app.Selection.Find.Wrap = 0
        app.Selection.Find.Text = "\{E(*)_"

def sub_condition():
    set_find(Text="\{E(*)_")
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        label_text = app.Selection.text[2:-1]
        #CONDITION JUDGE
        try:
            cond = 1 if conf.loc[label_text][0] == 1 else 0
        except Exception:
            cond = 0
        #NEED
        if cond == 1:
            oldStr = "\{E%s_(*)_E%s\}" % (label_text, label_text)
            newStr = "\\1"
            app.Selection.Find.Execute(oldStr, False, False, True, False, False, True, 1, False, newStr, 1)
        #DONT NEED
        elif cond == 0:
            app.Selection.Find.Text = "\{E%s_(*)_E%s\}" % (label_text, label_text)
            app.Selection.End = app.Selection.Start
            app.Selection.Find.Execute()
            app.selection.Font.Hidden = True

        app.Selection.Find.Text = "\{E(*)_"

# 9 Deal with conditions
def condition(mode):
    """
    cond from conf
    :param mode:
    :return:
    """
    level_1_condition()
    sub_condition()

    print("Func 9 Done!")

def Ftrans(path):

    target_doc = app.Documents.Open(path)
    doc1.Activate()
    set_find(Text="\{F(*)_")
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        label_text = app.Selection.text[2:-1]
        app.Selection.Find.Text = "\{F%s_(*)_F%s\}" % (label_text, label_text)
        app.Selection.End = app.Selection.Start
        app.Selection.Find.Execute()
        offset = len(label_text) + 3
        start = app.Selection.Start
        end = app.Selection.end
        doc1.Range(start+offset,end-offset).Copy()

        target_doc.Activate()
        app.Selection.Find.Text = "{F%s}" % label_text
        app.Selection.Find.MatchWildcards = False
        app.Selection.WholeStory()
        while app.Selection.Find.Execute():
            app.Selection.Paragraphs[0].Range.PasteAndFormat(16)

        doc1.Activate()
        app.Selection.Find.Text = "\{F(*)_"
        app.Selection.Find.MatchWildcards = True
        app.Selection.Start = app.Selection.End
        app.Selection.Find.Wrap = 0

    target_doc.Activate()
    target_doc.Save()
    target_doc.Close()
    doc1.Activate()

#10 Field Trans
def Ftrans_single(mode='conf'):
    """
    :param mode: input: input conf: conf file
           path: target doc path
    :return:
    """

    path = None
    if mode == 'input':
        path = input("请输入目的文件绝对路径：")
    elif mode == 'conf':
        path = conf.loc['10_single_path'][0]

    Ftrans(path)

    print("Func 10 Done!")


#11 Inner Field Trans
def Finner_copy(mode):
    """
    Inner Field Trans
    :return:
    """
    set_find(Text="\{F(*)_")
    app.Selection.WholeStory()
    while app.Selection.Find.Execute():
        label_text = app.Selection.text[2:-1]
        app.Selection.Find.Text = "\{F%s_(*)_F%s\}" % (label_text, label_text)
        app.Selection.End = app.Selection.Start
        app.Selection.Find.Execute()
        offset = len(label_text) + 3
        start = app.Selection.Start
        end = app.Selection.end
        doc1.Range(start+offset,end-offset).Copy()

        app.Selection.Find.Text = "{F%s}" % label_text
        app.Selection.Find.MatchWildcards = False
        app.Selection.Start = 0
        app.Selection.End = 0
        while app.Selection.Find.Execute():
            app.Selection.Paragraphs[0].Range.PasteAndFormat(16)

        app.Selection.Find.Text = "\{F(*)_"
        app.Selection.Find.MatchWildcards = True
        app.Selection.Start = end
        app.Selection.End = end
        app.Selection.Find.Wrap = 0

    print("Func 11 Done!")

#12 Batch Replace
def batch_replace(mode='conf'):
    """
    Batch Replace
    :param mode: input: input conf: conf file
           directory: Dir containing multiple fields
    :return:
    """

    directory = None
    if mode == 'input':
        directory = input("请输入批替换的目录路径：")
    elif mode == 'conf':
        directory = conf.loc['12_batch_dir'][0]

    docx_paths = fiFindByWildcard(os.path.join(directory, '*.docx'))
    for path in docx_paths:
        temp_doc = app.Documents.Open(path)
        temp_doc.Activate()
        replace_elements(temp_doc.Name)
        temp_doc.Save()
        temp_doc.Close()

    doc1.Activate()

    print("Func 12 Done!")

#13 Batch Trans
def batch_trans(mode = 'conf'):
    """
    Batch Trans
    :param mode: input: input conf: conf file
           directory: Dir containing multiple fields
    :return:
    """

    directory = None
    if mode == 'input':
        directory = input("请输入批转移的目录路径：")
    elif mode == 'conf':
        directory = conf.loc['13_batch_dir'][0]

    docx_paths = fiFindByWildcard(os.path.join(directory, '*.docx'))
    for path in docx_paths:
        Ftrans(path)

    doc1.Activate()

    print("Func 13 Done!")

#14 Change doc1:
def change_doc1(mode):
    path = input("请输入doc_source绝对路径：")
    global doc1
    doc1.Save()
    doc1.Close()
    doc1 = app.Documents.Open(path)
    doc1.Activate()
    print("Func 14 Done!")




#----------------------------------------Testing--------------------------------------

# a = doc1.paragraphs[24].range.highlightcolorindex
# b = doc1.paragraphs[22].range.font.colorindex
#insert col
# for i in range(1, doc1.tables[0].rows.count + 1):
#     doc1.tables[0].Cell(i, doc1.tables[0].columns.count).range.text = 'content_{}'.format(i)
# doc1.tables[0].Cell(1,1).range.text = '具体项目'

func_dict = {"1":remove_color,"2":hide_all_T,"3":copy_chosen_T,"4":restore,"5":replace_doc1,"6":remove_mark,
             "7":remove_condition_row,"8":delete_trigger_para,"9":condition,"10":Ftrans_single,"11":Finner_copy,"12":batch_replace,
             "13":batch_trans,"14":change_doc1}
while True:
    cmd = input("请输入模式和命令（或组合）：")

    if cmd == 'quit':
        sys.exit()
    else:
        Mode,cmd = cmd.split(':')[0],cmd.split(':')[1]
        Mode = 'input' if Mode == 'i' else 'conf' if Mode == 'c' else 'error'
        if Mode == 'error':
            print('Wrong Mode!')
            sys.exit()

        if re.match("^\d+$", cmd):
            func_dict[cmd](Mode)
        elif re.match("^(\d+;)+(\d+)$", cmd):
            for index in cmd.split(';'):
                func_dict[index](Mode)
        print('Task Done!')

# copy_chosen_T(chosen=2,t_num=4)
# hide_all_T(t_num=4)
# level_1_condition()
# sub_condition()
# replace_elements('main')
# remove_color('橙色')
# remove_condition_row(enum='-')
# remove_condition_row(enum='')

# #提交时
# remove_mark()
# #恢复模板状态
# restore()

# delete_trigger_para("{Trigger}")

# batch_replace(r'C:\Users\lihongyu\Desktop\testDir')
# Ftrans(doc2)


