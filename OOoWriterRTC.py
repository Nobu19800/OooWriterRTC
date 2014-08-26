# -*- coding: utf-8 -*-

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math


if os.name == 'posix':
    sys.path += ['/usr/lib/python2.6/dist-packages', '/usr/lib/python2.6/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['C:\\Python26\\lib\\site-packages', 'C:\\Python26\\lib\\site-packages\\rtctree\\rtmidl']

import time
import random
import commands
import RTC
import OpenRTM_aist

from OpenRTM_aist import CorbaNaming
from OpenRTM_aist import RTObject
from OpenRTM_aist import CorbaConsumer
from omniORB import CORBA
import CosNaming





import uno
import unohelper

from com.sun.star.awt.FontWeight import BOLD
from com.sun.star.awt.FontWeight import NORMAL
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt.FontSlant import NONE
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.beans import PropertyValue
from com.sun.star.table import TableBorder
from com.sun.star.text import TableColumnSeparator
from com.sun.star.text.HoriOrientation import NONE as HO_NONE


import OOoRTC





#comp_num = random.randint(1,3000)
imp_id = "OOoWriterControl"# + str(comp_num)







ooowritercontrol_spec = ["implementation_id", imp_id,
                  "type_name",         imp_id,
                  "description",       "Openoffice Writer Component",
                  "version",           "0.1",
                  "vendor",            "Miyamoto Nobuhiko",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  "conf.default.fontsize", "16",
                  #"conf.default.fontname", "ＭＳ 明朝",
                  "conf.default.Red", "0",
                  "conf.default.Blue", "0",
                  "conf.default.Green", "0",
                  "conf.default.Italic", "0",
                  "conf.default.Bold", "0",
                  "conf.default.Code", "utf-8",
                  "conf.__widget__.fontsize", "spin",
                  #"conf.__widget__.fontname", "radio",
                  "conf.__widget__.Red", "spin",
                  "conf.__widget__.Blue", "spin",
                  "conf.__widget__.Green", "spin",
                  "conf.__widget__.Italic", "radio",
                  "conf.__widget__.Bold", "radio",
                  "conf.__widget__.Code", "radio",
                  "conf.__constraints__.fontsize", "1<=x<=72",
                  #"conf.__constraints__.fontname", "(MS UI Gothic,MS ゴシック,MS Pゴシック,MS 明朝,MS P明朝,HG ゴシック E,HGP ゴシック E,HGS ゴシック E,HG ゴシック M,HGP ゴシック M,HGS ゴシック M,HG 正楷書体-PRO,HG 丸ゴシック M-PRO,HG 教科書体,HGP 教科書体,HGS 教科書体,HG 行書体,HGP 行書体,HGS 行書体,HG 創英プレゼンス EB,HGP 創英プレゼンス EB,HGS 創英プレゼンス EB,HG 創英角ゴシック UB,HGP 創英角ゴシック UB,HGS 創英角ゴシック UB,HG 創英角ポップ体,HGP 創英角ポップ体,HGS 創英角ポップ体,HG 明朝 B,HGP 明朝 B,HGS 明朝 B,HG 明朝 E,HGP 明朝 E,HGS 明朝 E,メイリオ)",
                  "conf.__constraints__.Red", "0<=x<=255",
                  "conf.__constraints__.Blue", "0<=x<=255",
                  "conf.__constraints__.Green", "0<=x<=255",
                  "conf.__constraints__.Italic", "(0,1)",
                  "conf.__constraints__.Bold", "(0,1)",
                  "conf.__constraints__.Code", "(utf-8,euc_jp,shift_jis)",
    ""
                  ""]

def SetCoding(m_str, m_code):
    if os.name == 'posix':
        if m_code == "utf-8":
            return m_str
        else:
            try:
                return m_str.decode(m_code).encode("utf-8")
            except:
                return ""
    elif os.name == 'nt':
        try:
            return m_str.decode(m_code).encode('cp932')
        except:
            return ""

def ResetCoding(m_str):
    if os.name == 'posix':
        return m_str.encode('utf-8')
    elif os.name == 'nt':
        return m_str.encode('cp932')








##
# OpenOffice Writerを操作するためのRTCのクラス
##

class OOoWriterControl(OpenRTM_aist.DataFlowComponentBase):
  def __init__(self, manager):
    OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
    
    
    self._d_m_word = RTC.TimedString(RTC.Time(0,0),0)
    self._m_wordIn = OpenRTM_aist.InPort("word", self._d_m_word)

    self._d_m_fontSize = RTC.TimedFloat(RTC.Time(0,0),0)
    self._m_fontSizeIn = OpenRTM_aist.InPort("fontSize", self._d_m_fontSize)

    self._d_m_fontName = RTC.TimedString(RTC.Time(0,0),0)
    self._m_fontNameIn = OpenRTM_aist.InPort("fontName", self._d_m_fontName)

    self._d_m_wsCharacter = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsCharacterIn = OpenRTM_aist.InPort("wsCharacter", self._d_m_wsCharacter)

    self._d_m_wsWord = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsWordIn = OpenRTM_aist.InPort("wsWord", self._d_m_wsWord)

    self._d_m_wsLine = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsLineIn = OpenRTM_aist.InPort("wsLine", self._d_m_wsLine)

    self._d_m_wsParagraph = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsParagraphIn = OpenRTM_aist.InPort("wsParagraph", self._d_m_wsParagraph)

    self._d_m_wsWindow = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsWindowIn = OpenRTM_aist.InPort("wsWindow", self._d_m_wsWindow)

    self._d_m_wsScreen = RTC.TimedShort(RTC.Time(0,0),0)
    self._m_wsScreenIn = OpenRTM_aist.InPort("wsScreen", self._d_m_wsScreen)

    self._d_m_color = RTC.TimedRGBColour(RTC.Time(0,0),RTC.RGBColour(0,0,0))
    self._m_colorIn = OpenRTM_aist.InPort("color", self._d_m_color)

    self._d_m_MovementType = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_MovementTypeIn = OpenRTM_aist.InPort("MovementType", self._d_m_MovementType)

    self._d_m_Italic = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_ItalicIn = OpenRTM_aist.InPort("Italic", self._d_m_Italic)

    self._d_m_Bold = RTC.TimedBoolean(RTC.Time(0,0),0)
    self._m_BoldIn = OpenRTM_aist.InPort("Bold", self._d_m_Bold)

    self._d_m_selWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_selWordOut = OpenRTM_aist.OutPort("selWord", self._d_m_selWord)

    self._d_m_copyWord = RTC.TimedString(RTC.Time(0,0),0)
    self._m_copyWordOut = OpenRTM_aist.OutPort("copyWord", self._d_m_copyWord)
    

    try:
      self.writer = OOoWriter()
    except NotOOoWtiterException:
      return

    self.fontSize = 16
    self.fontName = "ＭＳ 明朝"
    self.Bold = False
    self.Italic = False
    self.Red = 0
    self.Green = 0
    self.Blue = 0
    self.MovementType = False


    self.conf_fontSize = [16]
    self.conf_fontName = ["ＭＳ 明朝"]
    self.conf_Bold = [False]
    self.conf_Italic = [False]
    self.conf_Red = [0]
    self.conf_Green = [0]
    self.conf_Blue = [0]
    self.conf_Code = ["utf-8"]
    
    
    return

  ##
  # 実行周期を設定する関数
  ##
  def m_setRate(self, rate):
      m_ec = self.get_owned_contexts()
      m_ec[0].set_rate(rate)

  ##
  # 活性化するための関数
  ## 
  def m_activate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].activate_component(self._objref)

  ##
  # 不活性化するための関数
  ##
  def m_deactivate(self):
      m_ec = self.get_owned_contexts()
      m_ec[0].deactivate_component(self._objref)

  


  ##
  # 初期化処理用コールバック関数
  ##
  def onInitialize(self):
    
    OOoRTC.writer_comp = self

    self.addInPort("word",self._m_wordIn)
    self.addInPort("fontSize",self._m_fontSizeIn)
    self.addInPort("wsCharacter",self._m_wsCharacterIn)
    self.addInPort("wsWord",self._m_wsWordIn)
    self.addInPort("wsLine",self._m_wsLineIn)
    self.addInPort("wsParagraph",self._m_wsParagraphIn)
    self.addInPort("wsWindow",self._m_wsWindowIn)
    self.addInPort("wsScreen",self._m_wsScreenIn)
    self.addInPort("color",self._m_colorIn)
    self.addInPort("MovementType",self._m_MovementTypeIn)
    self.addInPort("Italic",self._m_ItalicIn)
    self.addInPort("Bold",self._m_BoldIn)
    self.addOutPort("selWord",self._m_selWordOut)
    self.addOutPort("copyWord",self._m_copyWordOut)

    self.bindParameter("fontsize", self.conf_fontSize, "16")
    #self.bindParameter("fontname", self.conf_fontName, "ＭＳ 明朝")
    self.bindParameter("Bold", self.conf_Bold, "False")
    self.bindParameter("Italic", self.conf_Italic, "False")
    self.bindParameter("Red", self.conf_Red, "0")
    self.bindParameter("Blue", self.conf_Green, "0")
    self.bindParameter("Green", self.conf_Blue, "0")
    self.bindParameter("Code", self.conf_Code, "utf-8")

    
    
    
    return RTC.RTC_OK

  ##
  # 文字書き込みの関数
  ##

  def SetWord(self, m_str):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      inp_str = SetCoding(m_str, self.conf_Code[0])
      cursor.setString(inp_str)
      
      cursor.CharHeight = self.fontSize
      cursor.CharHeightAsian = self.fontSize

      

      if self.Bold:
          cursor.CharWeight = BOLD
          cursor.CharWeightAsian = BOLD
      else:
          cursor.CharWeight = NORMAL
          cursor.CharWeightAsian = NORMAL
      if self.Italic:
          cursor.CharPosture = ITALIC
          cursor.CharPostureAsian = ITALIC
      else:
          cursor.CharPosture = NONE
          cursor.CharPostureAsian = NONE

      #cursor.CharStyleName = self.fontName
       
      
      

      cursor.goRight(len(inp_str),False)

      cursor.collapseToEnd()

  ##
  # カーソル位置の文字取得の関数
  ##

  def GetWord(self):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      try:
          out_str = ResetCoding(cursor.getString())
          return out_str
      except:
          return ""
       
      

  
      
      

  ##
  # 文字数移動する関数
  ##
  def MoveCharacter(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goRight(diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToEnd()
      else:
          cursor.goLeft(-diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToStart()
          
  ##
  # 単語数移動する関数
  ##
  def MoveWord(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextWord(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToEnd()
          else:
              cursor.gotoPreviousWord(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToStart()

  ##
  # 行数移動する関数
  ##
  def MoveLine(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      if diff > 0:
          cursor.goDown(diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToEnd()
      else:
          cursor.goUp(-diff,self.MovementType)
          if self.MovementType == False:
              cursor.collapseToStart()

  ##
  # 段落数移動する関数
  ##
  def MoveParagraph(self, diff):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      for i in range(0, diff):
          if diff > 0:
              cursor.gotoNextParagraph(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToEnd()
          else:
              cursor.gotoPreviousParagraph(self.MovementType)
              if self.MovementType == False:
                  cursor.collapseToStart()


  ##
  # 活性化処理用コールバック関数
  ##
  
  def onActivated(self, ec_id):
    self.fontSize = float(self.conf_fontSize[0])
    self.fontName = self.conf_fontName[0]
    if int(self.conf_Bold[0]) == 0:
        self.Bold = False
    else:
        self.Bold = True
    if int(self.conf_Italic[0]) == 0:
        self.Italic = False
    else:
        self.Italic = True
    self.Red = int(self.conf_Red[0])
    self.Green = int(self.conf_Green[0])
    self.Blue = int(self.conf_Blue[0])
    
    return RTC.RTC_OK   

  ##
  # 周期処理用コールバック関数
  ##
  
  def onExecute(self, ec_id):
    

    if self._m_fontSizeIn.isNew():
        data = self._m_fontSizeIn.read()
        self.fontSize = data.data

    if self._m_wsCharacterIn.isNew():
        data = self._m_wsCharacterIn.read()
        self.MoveCharacter(data.data)

    if self._m_wsWordIn.isNew():
        data = self._m_wsWordIn.read()
        self.MoveWord(data.data)

    if self._m_wsLineIn.isNew():
        data = self._m_wsLineIn.read()
        self.MoveLine(data.data)

    if self._m_wsParagraphIn.isNew():
        data = self._m_wsParagraphIn.read()
        self.MoveParagraph(data.data)

    if self._m_wsWindowIn.isNew():
        data = self._m_wsWindowIn.read()
        pass


    if self._m_wsScreenIn.isNew():
        data = self._m_wsScreenIn.read()
        pass

    if self._m_colorIn.isNew():
        data = self._m_colorIn.read()
        self.Red = data.data.r*255
        self.Green = data.data.g*255
        self.Blue = data.data.b*255

    if self._m_ItalicIn.isNew():
        data = self._m_ItalicIn.read()
        self.Italic = data.data
    
    

    if self._m_BoldIn.isNew():
        data = self._m_BoldIn.read()
        self.Bold = data.data

    if self._m_MovementTypeIn.isNew():
        data = self._m_MovementTypeIn.read()
        self.MovementType = data.data

    if self._m_wordIn.isNew():
        data = self._m_wordIn.read()
        self.SetWord(data.data)

    
    self._d_m_selWord.data = str(self.GetWord())
    self._m_selWordOut.write()
        

    return RTC.RTC_OK

  ##
  # 終了処理用コールバック関数
  ##
  
  def on_shutdown(self, ec_id):
      OOoRTC.writer_comp = None
      return RTC.RTC_OK



##
# コンポーネントを活性化してWriterの操作を開始する関数
##

def Start():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_activate()

##
# コンポーネントを不活性化してWriterの操作を終了する関数
##

def Stop():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.m_deactivate()


##
# コンポーネントの実行周期を設定する関数
##

def Set_Rate():
    pass
    """if OOoRTC.writer_comp:
      try:
        writer = OOoDraw()
      except NotOOoWriterException:
          return

      oWriterPages = draw.drawpages
      for i in range(0, oDrawPages.Count):
        oDrawPage = oDrawPages.getByIndex(i)
        forms = oDrawPage.getForms()
        for j in range(0, forms.Count):
          st_control = oDrawPage.getForms().getByIndex(j).getByName('Rate')
          if st_control:
            try:
              text = float(st_control.Text)
            except:
               return
              
            OOoRTC.draw_comp.m_setRate(text)"""
      
      

      
        
        
      
      

      



      
  



##
#RTCをマネージャに登録する関数
##
def OOoWriterControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooowritercontrol_spec)
  manager.registerFactory(profile,
                          OOoWriterControl,
                          OpenRTM_aist.Delete)


def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoWriterControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# RTC起動の関数
##

def createOOoWriterComp():
                        
    
    if OOoRTC.mgr == None:
      OOoRTC.mgr = OpenRTM_aist.Manager.init(['OOoWriter.py'])
      OOoRTC.mgr.setModuleInitProc(MyModuleInit)
      OOoRTC.mgr.activateManager()
      OOoRTC.mgr.runManager(True)
    else:
      MyModuleInit(OOoRTC.mgr)
      
          

    try:
      writer = OOoWriter()
    except NotOOoWriterException:
      return

    
    MyMsgBox('',SetCoding('RTCを起動しました','utf-8'))


    
    
    return None




##
# メッセージボックス表示の関数
# title：ウインドウのタイトル
# message：表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
##

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
    msgbox.execute()
    msgbox.dispose()





##
# OpenOffice Writerを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部を改変
##

class OOoWriter(Bridge):
  def __init__(self):
    Bridge.__init__(self)
    if not self._document.supportsService('com.sun.star.text.TextDocument'):
      self.run_errordialog(title='エラー', message='このマクロはOpenOffice.org Writerの中で実行してください')
      raise NotOOoWriterException()
    self.__current_controller = self._document.CurrentController
    
  @property
  def document(self): return self._document
  



    


g_exportedScripts = (createOOoWriterComp, Start, Stop, Set_Rate)
