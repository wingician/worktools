# -*- coding: SHIFT_JIS -*-

#Boa:Frame:Frame2


import wx
import os.path
import time
import sys, os
reload(sys)
sys.setdefaultencoding('SHIFT_JIS')

import win32com.client
import win32api
import string
import re
import xlrd
import wx

import Frame2

modules ={'Frame2': [1, 'Main frame of Application', u'Frame2.py']}

class BoaApp(wx.App):
    def OnInit(self):
        self.main = Frame2.create(None)
        self.main.Show()
        self.SetTopWindow(self.main)
        return True

def main():
    application = BoaApp(0)
    application.MainLoop()

if __name__ == '__main__':
    main()



def create(parent):
    return Frame2(parent)

[wxID_FRAME2, wxID_FRAME2BUTTON1, wxID_FRAME2DIRPICKERCTRL1, 
 wxID_FRAME2DIRPICKERCTRL2, wxID_FRAME2RADIOBUTTON1, wxID_FRAME2RADIOBUTTON2, 
 wxID_FRAME2RADIOBUTTON3, wxID_FRAME2RADIOBUTTON4, wxID_FRAME2RADIOBUTTON5, 
 wxID_FRAME2STATICTEXT1, wxID_FRAME2STATICTEXT2, wxID_FRAME2STATICTEXT3, 
 wxID_FRAME2STATICTEXT4, wxID_FRAME2STATICTEXT5, wxID_FRAME2TEXTCTRL1, 
] = [wx.NewId() for _init_ctrls in range(15)]

class Frame2(wx.Frame):
    
    def __init__(self, parent):
        self._init_ctrls(parent)

    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME2, name='', parent=prnt,
              pos=wx.Point(423, 280), size=wx.Size(406, 352),
              style=wx.DEFAULT_FRAME_STYLE, title=u'Tiny Grep')
        self.SetClientSize(wx.Size(390, 314))
        self.SetBackgroundColour(wx.Colour(235, 235, 235))
        self.SetBackgroundStyle(wx.BG_STYLE_COLOUR)
        self.SetAutoLayout(True)

        self.textCtrl1 = wx.TextCtrl(id=wxID_FRAME2TEXTCTRL1, name='textCtrl1',
              parent=self, pos=wx.Point(26, 32), size=wx.Size(238, 144),
              style=wx.TE_MULTILINE, value='')

        self.staticText1 = wx.StaticText(id=wxID_FRAME2STATICTEXT1,
              label=u'\u6587\u5b57\u5217\uff08\u8907\u6570\u306e\u5834\u5408\u3001\u6539\u884c\u5165\u529b\u304f\u3060\u3055\u3044\u3002\u7a7a\u884c\u7981\u6b62\uff09',
              name='staticText1', parent=self, pos=wx.Point(24, 8),
              size=wx.Size(288, 19), style=0)

        self.button1 = wx.Button(id=wxID_FRAME2BUTTON1, label=u'\u691c\u7d22',
              name='button1', parent=self, pos=wx.Point(272, 216),
              size=wx.Size(88, 24), style=0)
        self.button1.Bind(wx.EVT_BUTTON, self.OnButton1Button,
              id=wxID_FRAME2BUTTON1)

        self.staticText2 = wx.StaticText(id=wxID_FRAME2STATICTEXT2,
              label=u'\u691c\u7d22\u5834\u6240\uff1a', name='staticText2',
              parent=self, pos=wx.Point(24, 192), size=wx.Size(60, 19),
              style=0)

        self.dirPickerCtrl1 = wx.DirPickerCtrl(id=wxID_FRAME2DIRPICKERCTRL1,
              message='Select a folder', name='dirPickerCtrl1', parent=self,
              path='', pos=wx.Point(24, 216),
              style=wx.DIRP_CHANGE_DIR|wx.DIRP_USE_TEXTCTRL)

        self.radioButton1 = wx.RadioButton(id=wxID_FRAME2RADIOBUTTON1,
              label=u'ALL(No Excel)', name='radioButton1', parent=self,
              pos=wx.Point(280, 40), size=wx.Size(105, 18), style=0)
        self.radioButton1.SetValue(True)

        self.radioButton2 = wx.RadioButton(id=wxID_FRAME2RADIOBUTTON2,
              label=u'JSP Only', name='radioButton2', parent=self,
              pos=wx.Point(280, 72), size=wx.Size(105, 18), style=0)
        self.radioButton2.SetValue(False)

        self.radioButton3 = wx.RadioButton(id=wxID_FRAME2RADIOBUTTON3,
              label=u'Cobol Only', name='radioButton3', parent=self,
              pos=wx.Point(280, 136), size=wx.Size(105, 18), style=0)
        self.radioButton3.SetValue(False)

        self.radioButton4 = wx.RadioButton(id=wxID_FRAME2RADIOBUTTON4,
              label=u'SHL Only', name='radioButton4', parent=self,
              pos=wx.Point(280, 104), size=wx.Size(105, 18), style=0)
        self.radioButton4.SetValue(False)

        self.staticText3 = wx.StaticText(id=wxID_FRAME2STATICTEXT3,
              label=u'by scs\u3000v0.2', name=u'staticText3', parent=self,
              pos=wx.Point(296, 288), size=wx.Size(65, 16), style=0)

        self.staticText4 = wx.StaticText(id=wxID_FRAME2STATICTEXT4,
              label=u'\u7d50\u679c\u51fa\u529b\uff1a(grepOutPut.csv)',
              name='staticText4', parent=self, pos=wx.Point(24, 248),
              size=wx.Size(208, 16), style=0)

        self.dirPickerCtrl2 = wx.DirPickerCtrl(id=wxID_FRAME2DIRPICKERCTRL2,
              message='Select a folder', name='dirPickerCtrl2', parent=self,
              path='', pos=wx.Point(24, 272), style=wx.DIRP_DEFAULT_STYLE)

        self.staticText5 = wx.StaticText(id=wxID_FRAME2STATICTEXT5,
              label=u'Ready', name='staticText5', parent=self, pos=wx.Point(280,
              256), size=wx.Size(96, 16), style=0)

        self.radioButton5 = wx.RadioButton(id=wxID_FRAME2RADIOBUTTON5,
              label=u'Excel', name=u'radioButton5', parent=self,
              pos=wx.Point(280, 168), size=wx.Size(105, 18), style=0)
        self.radioButton5.SetValue(False)
        self.radioButton5.Bind(wx.EVT_RADIOBUTTON,
              self.OnRadioButton5Radiobutton, id=wxID_FRAME2RADIOBUTTON5)

    def OnAbout(self, event):
        dialog = wx.MessageDialog(self, 'A Tiny grep tools\n'
            'in wxPython', 'About Tiny Grep', wx.OK)
        dialog.ShowModal()
        dialog.Destroy()

    def OnExit(self, event):
        self.Close()  # Close the main window.
                
        
    def OnButton1Button(self, event):
        
        self.staticText5.SetLabel(u'検索中...')
        outputdir =  self.dirPickerCtrl2.GetPath()
        
        outfileName = 'grepOutPut.csv'
        outfile = os.path.join(outputdir,outfileName)
        fout = open(outfile, 'w')
        
        #let grep key word in file
        keywords = self.textCtrl1.GetValue()
        keywordName = 'grepKeyWord.txt'
        keywordFile = os.path.join(outputdir,keywordName)
        fkeyword = open(keywordFile, 'w')      #open keywordfile 
        fkeyword.write(keywords)
        fkeyword.close()
        
        #and open key word file (because jp encoding problom.)
        keywordlist = open(keywordFile,'r') 
        errorFile = open('errListFile.txt','w')
        
        

        
        #search target path
        searchdir = self.dirPickerCtrl1.GetPath()
        
        semi = ','
        #result in excel file
        #Application = win32com.client.Dispatch("Excel.Application")  #out put file
        #Application.Visible = 1
        #WorkBook = Application.Workbooks.Add()
        #Base = WorkBook.ActiveSheet
        #Base.Cells(1,1).Value =  u'交番'
        #Base.Cells(1,2).Value =  u'検索キー'
        #Base.Cells(1,3).Value =  u'プログラムID'
        #Base.Cells(1,4).Value =  u'パス'
        #Base.Cells(1,5).Value =  u'ソース抜粋'    
        outPutTitle = u'項番' + ',' + u'検索キー' + ',' + u'プログラムID' + ',' + u'パス' + ',' + u'ソース抜粋（turn カンマ to セミコロン）' + '\n'
        fout.write(outPutTitle)
        
        count = 0
        errorFlg = 0
        
        #search file types.
        if self.radioButton2.Value:
            grepType0 = 'jsp'
            grepType1 = 'inc'
    
        if self.radioButton3.Value:
            grepType0 = 'pco'
            grepType1 = 'cpy'
         
                
        if self.radioButton4.Value:
            grepType0 = 'csh'
            grepType1 = 'sql'

        if self.radioButton5.Value:
            grepType0 = 'xls'
            grepType1 = 'xlsx'
                                        
    
        for gkey in keywordlist:
            if gkey !='':
                keywordForoutput = gkey
                keywordForoutput = keywordForoutput.strip()
                gkey = gkey.lower()
                gkey = gkey.strip()
                keyword = gkey
                for (dirname, dirs, files) in os.walk(searchdir):
                    for filename in files:
                        if self.radioButton1.Value:
                            thefile = os.path.join(dirname,filename)
                            in_file = open(thefile,'r')
                            for line in in_file:
                                lineout = "'" + line
                                line = line.lower()
                                line = line.strip()
                                #if re.search(gkey, line) :
                                testFlag = 1
                                #try:
                                if testFlag == 1:
                                    if line.find(gkey) != -1:
                                        count = count + 1
                                        lineoutWithout = str(lineout).replace(',',';')
                                        f_output = str(count) + semi + str(keywordForoutput) + semi + str(filename) + semi + str(dirname) + semi +  lineoutWithout 
                                        fout.write(f_output)
                                        #Base.Cells(count+1,1).Value = "'" + str(count)
                                        #Base.Cells(count+1,2).Value = "'" + keywordForoutput
                                        #Base.Cells(count+1,3).Value = "'" + str(filename)
                                        #Base.Cells(count+1,4).Value = "'" + str(dirname)
                                        #Base.Cells(count+1,5).Value = "'" + lineout
                                else:
                                    errorMSG = 'error in search ' + str(in_file) + '\n'
                                    errorFile.write(errorMSG)
                                    errorFlg = 1
                            in_file.close()
                        elif self.radioButton5.Value:
                            #excel grep
                            if filename.endswith(grepType0) :
                                thefile = os.path.join(dirname,filename)
                                workbook = xlrd.open_workbook(thefile)
                                worksheets = workbook.sheet_names()
                                for worksheet_name in worksheets:
                                    worksheet = workbook.sheet_by_name(worksheet_name)
    
                                    num_rows = worksheet.nrows
                                    num_cols = worksheet.ncols
                                    for rown in range(num_rows):
                                        
                                        for coln in range(num_cols):
                                            cell = worksheet.cell_value(rown,coln)
                                            cell_type = worksheet.cell_type(rown,coln)
                                            if cell_type == 1:
                                                
                                                cell = cell.encode('SHIFT_JIS','ignore')
                                                cellString = str(cell)
                                                lineoutWithout = "'" + cellString + "'"
                                                lineoutWithout = lineoutWithout.replace('\n',';')
                                                lineoutWithout = lineoutWithout.replace(',',';')
                                                cellString = cellString.lower()
                                                cellString = cellString.strip()
                                                if cellString.find(gkey) != -1:
                                                    count = count + 1
                                                    
                                                    f_output = str(count) + semi + str(keywordForoutput) + semi + str(filename) + semi + str(dirname) + semi +  lineoutWithout + semi + str(worksheet_name) + '\n'
                                                    fout.write(f_output)
                                #thefile.close()
                            else:
                                if filename.endswith(grepType1) :
                                    thefile = os.path.join(dirname,filename)
                                    workbook = xlrd.open_workbook(thefile)
                                    worksheets = workbook.sheet_names()
                                    for worksheet_name in worksheets:
                                        worksheet = workbook.sheet_by_name(worksheet_name)
    
                                        num_rows = worksheet.nrows
                                        num_cols = worksheet.ncols
                                        for rown in range(num_rows):
                                        
                                            for coln in range(num_cols):
                                                cell = worksheet.cell_value(rown,coln)
                                                cell_type = worksheet.cell_type(rown,coln)
                                                if cell_type == 1:
                                                
                                                    cell = cell.encode('SHIFT_JIS','ignore')
                                                    cellString = str(cell)
                                                    lineoutWithout = "'" + cellString + "'"
                                                    lineoutWithout = lineoutWithout.replace('\n',';')
                                                    lineoutWithout = lineoutWithout.replace(',',';')
                                                    cellString = cellString.lower()
                                                    cellString = cellString.strip()
                                                    if cellString.find(gkey) != -1:
                                                        count = count + 1
                                                    
                                                        f_output = str(count) + semi + str(keywordForoutput) + semi + str(filename) + semi + str(dirname) + semi +  lineoutWithout + semi + str(worksheet_name) + '\n'
                                                        fout.write(f_output)                               
                                #thefile.close()
                        else:
                            if filename.endswith(grepType0) :
                                thefile = os.path.join(dirname,filename)
                                in_file = open(thefile,'r')
                                for line in in_file:
                                    lineout = "'" + line
                                    line = line.lower()
                                    line = line.strip()
                                    testFlag = 1
                                    if testFlag == 1:
                                        if line.find(gkey) != -1:
                                            count = count + 1
                                            lineoutWithout = str(lineout).replace(',',';')
                                            f_output = str(count) + semi + str(keywordForoutput) + semi + str(filename) + semi + str(dirname) + semi +  lineoutWithout 
                                            fout.write(f_output)
                                    else:
                                        errorMSG = 'error in search ' + str(in_file) + '\n'
                                        errorFile.write(errorMSG)
                                        errorFlg = 1
                                in_file.close()
                            else:
                                if filename.endswith(grepType1) :
                                    thefile = os.path.join(dirname,filename)
                                    in_file = open(thefile,'r')
                                    for line in in_file:
                                        lineout = "'" + line
                                        line = line.lower()
                                        line = line.strip()
                                        
                                        testFlag = 1
                                        if testFlag == 1:
                                            if line.find(gkey) != -1:
                                                count = count + 1
                                                lineoutWithout = str(lineout).replace(',',';')
                                                f_output = str(count) + semi + str(keywordForoutput) + semi + str(filename) + semi + str(dirname) + semi +  lineoutWithout 
                                                fout.write(f_output)
                                                
                                        else:
                                            errorMSG = 'error in search ' + str(in_file) + '\n'
                                            errorFile.write(errorMSG)
                                            errorFlg = 1
                                    in_file.close()
        
        errorFile.close()       
        if   errorFlg == 0:                                       
            dialog = wx.MessageDialog(self, u'検索完了!\n 結果： grepOutPut.csv.'
                '', 'Tiny Grep ', wx.OK)
            dialog.ShowModal()
            dialog.Destroy()
        else:
            dialog = wx.MessageDialog(self, 'Grep Completed WITH SOME ERRORS!\n Show Results in Excel File.\n'
                'Error in errListFile.txt', 'Tiny Grep ^-^', wx.OK)
            dialog.ShowModal()
            dialog.Destroy()
            

        fout.close()
        keywordlist.close()
        os.remove(keywordFile)
        self.staticText5.SetLabel(u'Ready')
        pass

    def OndirPickerCtrl1Button(self, event):
        event.Skip()

    def OnRadioButton5Radiobutton(self, event):
        event.Skip()
      
  
