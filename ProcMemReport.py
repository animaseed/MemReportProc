import codecs
import os
import xlsxwriter

WORK_PATH = os.getcwd()


class TexInfo:
    def __init__(self):
        self.WxH = None
        self.Size = None
        self.Format = None
        self.LODGroup = None
        self.Name = None
        self.Streaming = None
        self.UnknownRef = None
        self.VT = None
        self.UsageCount = None
        self.NumMips = None
        self.Uncompressed = None
        self.FolderPath = None


def addWorkSheetTitle(ws):
    titleList = ['名称', '大小（KB）', '尺寸', '组别', '使用个数', 'Mip数', 'Format', 'Streaming', 'UnknownRef', 'VT', 'Uncompressed']
    for i in range(len(titleList)):
        title = titleList[i]
        ws.write(0, i, title)


def proc1MemReportFile(fileName):
    filePath = os.path.join(WORK_PATH, fileName)
    with codecs.open(filePath, 'r', 'utf-8') as f:
        lines = [x.strip() for x in f.readlines()]
    allBegin = -1
    allEnd = -1
    for i in range(len(lines)):
        line = lines[i]
        if line == 'MemReport: Begin command "ListTextures"':
            allBegin = i
        if line == 'MemReport: End command "ListTextures"':
            allEnd = i
    texGroupMap = {}
    folderPathMap = {}
    allInfoList = []
    for i in range(allBegin + 3, allEnd):
        line = lines[i]
        if line.startswith('Total size:'):
            break
        texInfo = TexInfo()
        infos = line.split(',')
        infos = [x.strip() for x in infos]
        wxhAndSize = infos[2]
        texInfo.WxH = wxhAndSize[:wxhAndSize.find('(')]
        texInfo.Size = int(wxhAndSize[wxhAndSize.find('(') + 1: wxhAndSize.find(')')][:-3])
        texInfo.Format = infos[3]
        texInfo.LODGroup = infos[4]
        Name = infos[5]
        Name = Name[:Name.rfind('.')]
        texInfo.Name = Name
        texInfo.Streaming = infos[6]
        texInfo.UnknownRef = infos[7]
        texInfo.VT = infos[8]
        texInfo.UsageCount = infos[9]
        texInfo.NumMips = infos[10]
        texInfo.Uncompressed = infos[11]
        artsPos = texInfo.Name.find('/Arts/')
        if artsPos == -1:
            texInfo.FolderPath = 'Other'
        else:
            folderPathStartPos = artsPos + 6
            folderPathEndPos = texInfo.Name.find('/', folderPathStartPos)
            texInfo.FolderPath = texInfo.Name[folderPathStartPos:folderPathEndPos]
        if texGroupMap.get(texInfo.LODGroup) is None:
            texGroupMap[texInfo.LODGroup] = []
        texGroupMap[texInfo.LODGroup].append(texInfo)
        if folderPathMap.get(texInfo.FolderPath) is None:
            folderPathMap[texInfo.FolderPath] = []
        folderPathMap[texInfo.FolderPath].append(texInfo)
        allInfoList.append(texInfo)
    page0Content = []
    for k, v in texGroupMap.items():
        totalSize = 0
        totalNum = 0
        for texInfo in v:
            totalNum = totalNum + 1
            totalSize = totalSize + texInfo.Size
        page0Content.append([k, str(totalSize), str(totalNum)])
    for k, v in folderPathMap.items():
        totalSize = 0
        totalNum = 0
        for texInfo in v:
            totalNum = totalNum + 1
            totalSize = totalSize + texInfo.Size
        page0Content.append([k, str(totalSize), str(totalNum)])
    excelFileName = fileName[:fileName.rfind('.')] + '.xlsx'
    wb = xlsxwriter.Workbook(os.path.join(WORK_PATH, excelFileName))
    ws = wb.add_worksheet('Total')
    ws.write(0, 0, '组别')
    ws.write(0, 1, '总内存（KB）')
    ws.write(0, 2, '数量')
    for i in range(len(page0Content)):
        info = page0Content[i]
        for j in range(len(info)):
            ws.write(i + 1, j, info[j])
    for k, v in texGroupMap.items():
        ws = wb.add_worksheet(k)
        addWorkSheetTitle(ws)
        for i in range(len(v)):
            txtInfo = v[i]
            ws.write(i + 1, 0, txtInfo.Name)
            ws.write(i + 1, 1, txtInfo.Size)
            ws.write(i + 1, 2, txtInfo.WxH)
            ws.write(i + 1, 3, txtInfo.LODGroup)
            ws.write(i + 1, 4, txtInfo.UsageCount)
            ws.write(i + 1, 5, txtInfo.NumMips)
            ws.write(i + 1, 6, txtInfo.Format)
            ws.write(i + 1, 7, txtInfo.Streaming)
            ws.write(i + 1, 8, txtInfo.UnknownRef)
            ws.write(i + 1, 9, txtInfo.VT)
            ws.write(i + 1, 10, txtInfo.Uncompressed)
    for k, v in folderPathMap.items():
        ws = wb.add_worksheet(k)
        addWorkSheetTitle(ws)
        for i in range(len(v)):
            txtInfo = v[i]
            ws.write(i + 1, 0, txtInfo.Name)
            ws.write(i + 1, 1, txtInfo.Size)
            ws.write(i + 1, 2, txtInfo.WxH)
            ws.write(i + 1, 3, txtInfo.LODGroup)
            ws.write(i + 1, 4, txtInfo.UsageCount)
            ws.write(i + 1, 5, txtInfo.NumMips)
            ws.write(i + 1, 6, txtInfo.Format)
            ws.write(i + 1, 7, txtInfo.Streaming)
            ws.write(i + 1, 8, txtInfo.UnknownRef)
            ws.write(i + 1, 9, txtInfo.VT)
            ws.write(i + 1, 10, txtInfo.Uncompressed)
    ws = wb.add_worksheet('All')
    addWorkSheetTitle(ws)
    for i in range(len(allInfoList)):
        txtInfo = allInfoList[i]
        ws.write(i + 1, 0, txtInfo.Name)
        ws.write(i + 1, 1, txtInfo.Size)
        ws.write(i + 1, 2, txtInfo.WxH)
        ws.write(i + 1, 3, txtInfo.LODGroup)
        ws.write(i + 1, 4, txtInfo.UsageCount)
        ws.write(i + 1, 5, txtInfo.NumMips)
        ws.write(i + 1, 6, txtInfo.Format)
        ws.write(i + 1, 7, txtInfo.Streaming)
        ws.write(i + 1, 8, txtInfo.UnknownRef)
        ws.write(i + 1, 9, txtInfo.VT)
        ws.write(i + 1, 10, txtInfo.Uncompressed)
    wb.close()


def run():
    fileNameList = os.listdir(WORK_PATH)
    for fileName in fileNameList:
        if os.path.isfile(os.path.join(WORK_PATH, fileName)) and fileName.endswith('.memreport'):
            proc1MemReportFile(fileName)


if __name__ == "__main__":
    run()
    # print(os.getcwd())
