#coding:utf8
from docx import *
import sys
import wx
import optparse
import os

def getWordText(path):
    os.system("xdoc2txt.exe -f %s" % path)
    textFile = os.path.splitext(path)[0] + ".txt"
    r = open(textFile,"r")
    ret = r.read()
    r.close()
    os.remove(textFile)
    return ret

if __name__ == '__main__':        
    app = wx.App(False)
    frame = wx.Frame(None,-1,"Word of Death",size=(600,800))
    panel = wx.Panel(frame,-1)
    panel.SetBackgroundColour("#AFAFAF")
    textCtrl = wx.TextCtrl(panel,-1,style=wx.TE_MULTILINE | wx.TE_RICH2)
    layout = wx.BoxSizer(wx.VERTICAL)
    layout.Add(textCtrl,proportion=1,flag=wx.EXPAND)
    panel.SetSizer(layout)
    panel.SetAutoLayout(True)
    frame.Show()
    oldFileList = []
    
    parser = optparse.OptionParser(usage="%prog new.docx old1.docx ... oldn.docx",version="%prog 0.1")
    parser.add_option("-a","--all",dest = "all")
    (opts,args) = parser.parse_args()
    if not opts.all:
        if len(args) == 1:
            print u"--allで全ワードファイルの差分取得"
            sys.exit(-1)
        for i in range(2,len(sys.argv)):
            oldFileList.append(sys.argv[i])
        newparatextlist = getWordText(sys.argv[1])
        oldparatextlistSet = [getWordText(oldFileList[x]) for x in range(len(oldFileList))]
    else:
        import glob
        docList = glob.glob("*.doc*")
        fTimes = []
        fileTimeDict = dict()
        for f in docList:
            fTimes.append(os.path.getmtime(f))
            fileTimeDict.update({os.path.getmtime(f):f})
        fTimes.sort()
        fTimes.reverse()

        print fileTimeDict

        newparatextlist = getWordText(fileTimeDict[fTimes[0]])
        oldparatextlistSet = [getWordText(fileTimeDict[fTimes[x]]) for x in range(1,len(fTimes))]


    

    for line in newparatextlist:
        findFlag = False
        for x in range(len(oldparatextlistSet)):
            if line in oldparatextlistSet[x]:
                findFlag = True
                if x == 0 :
                    textCtrl.SetDefaultStyle(wx.TextAttr("red"))
                elif x == 1:
                    textCtrl.SetDefaultStyle(wx.TextAttr("#FFA500"))
                elif x == 2:
                    textCtrl.SetDefaultStyle(wx.TextAttr("#00800"))
                elif x == 3:
                    textCtrl.SetDefaultStyle(wx.TextAttr("#7cfc00"))
                elif x == 4:
                    textCtrl.SetDefaultStyle(wx.TextAttr("#0000ff"))
                else:
                    textCtrl.SetDefaultStyle(wx.TextAttr("#4169e1"))

                textCtrl.AppendText(line+"\n")
                textCtrl.Refresh()
                textCtrl.SetDefaultStyle(wx.TextAttr("black"))
                break
        if not findFlag:
            textCtrl.AppendText(line+"\n")
            textCtrl.Refresh()

    
    app.MainLoop()
