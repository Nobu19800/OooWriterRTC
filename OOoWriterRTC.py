# -*- coding: utf-8 -*-



##
#   @file OOoWriterRTC.py
#   @brief OOoWriterControl Component

import optparse
import sys,os,platform
import re
import time
import random
import commands
import math

import os.path

from os.path import expanduser
sv = sys.version_info


if os.name == 'posix':
    home = expanduser("~")
    sys.path += [home+'/OOoRTC', home+'/OOoRTC/WriterIDL', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages', '/usr/lib/python2.' + str(sv[1]) + '/dist-packages/rtctree/rtmidl']
elif os.name == 'nt':
    sys.path += ['.\\OOoRTC', '.\\OOoRTC\\WriterIDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages', 'C:\\Python2' + str(sv[1]) + '\\Lib\\site-packages\\OpenRTM_aist\\RTM_IDL', 'C:\\Python2' + str(sv[1]) + '\\lib\\site-packages\\rtctree\\rtmidl']
    



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
from com.sun.star.awt.FontWeight import NORMAL as FW_NORMAL
from com.sun.star.awt.FontSlant import ITALIC
from com.sun.star.awt.FontSlant import NONE as FS_NONE
from com.sun.star.awt.FontUnderline import SINGLE as FU_SINGLE
from com.sun.star.awt.FontUnderline import NONE as FU_NONE
from com.sun.star.awt.FontStrikeout import SINGLE as FST_SINGLE
from com.sun.star.awt.FontStrikeout import NONE as FST_NONE
#from com.sun.star.awt.FontEmphasis import NONE as FE_NONE
#from com.sun.star.awt.FontRelief import ENGRAVED
#from com.sun.star.awt.FontRelief import NONE as FR_NONE
from com.sun.star.awt import XActionListener

from com.sun.star.script.provider import XScriptContext

from com.sun.star.beans import PropertyValue
from com.sun.star.table import TableBorder
from com.sun.star.text import TableColumnSeparator
from com.sun.star.text.HoriOrientation import NONE as HO_NONE


import OOoRTC
from WriterControl import *



import Writer_idl

from omniORB import PortableServer
import Writer, Writer__POA




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
                  "conf.default.Char_Red", "0",
                  "conf.default.Char_Blue", "0",
                  "conf.default.Char_Green", "0",
                  "conf.default.Italic", "0",
                  "conf.default.Bold", "0",
                  "conf.default.Underline", "0",
                  "conf.default.Shadow", "0",
                  "conf.default.Strikeout", "0",
                  "conf.default.Contoured", "0",
                  "conf.default.Emphasis", "0",
                  "conf.default.Back_Red", "255",
                  "conf.default.Back_Blue", "255",
                  "conf.default.Back_Green", "255",
                  "conf.default.Code", "utf-8",
                  "conf.__widget__.fontsize", "spin",
                  #"conf.__widget__.fontname", "radio",
                  "conf.__widget__.Char_Red", "spin",
                  "conf.__widget__.Char_Blue", "spin",
                  "conf.__widget__.Char_Green", "spin",
                  "conf.__widget__.Italic", "radio",
                  "conf.__widget__.Bold", "radio",
                  "conf.__widget__.Underline", "radio",
                  "conf.__widget__.Shadow", "radio",
                  "conf.__widget__.Strikeout", "radio",
                  "conf.__widget__.Contoured", "radio",
                  "conf.__widget__.Emphasis", "radio",
                  "conf.__widget__.Back_Red", "spin",
                  "conf.__widget__.Back_Blue", "spin",
                  "conf.__widget__.Back_Green", "spin",
                  "conf.__widget__.Code", "radio",
                  "conf.__constraints__.fontsize", "1<=x<=72",
                  #"conf.__constraints__.fontname", "(MS UI Gothic,MS ゴシック,MS Pゴシック,MS 明朝,MS P明朝,HG ゴシック E,HGP ゴシック E,HGS ゴシック E,HG ゴシック M,HGP ゴシック M,HGS ゴシック M,HG 正楷書体-PRO,HG 丸ゴシック M-PRO,HG 教科書体,HGP 教科書体,HGS 教科書体,HG 行書体,HGP 行書体,HGS 行書体,HG 創英プレゼンス EB,HGP 創英プレゼンス EB,HGS 創英プレゼンス EB,HG 創英角ゴシック UB,HGP 創英角ゴシック UB,HGS 創英角ゴシック UB,HG 創英角ポップ体,HGP 創英角ポップ体,HGS 創英角ポップ体,HG 明朝 B,HGP 明朝 B,HGS 明朝 B,HG 明朝 E,HGP 明朝 E,HGS 明朝 E,メイリオ)",
                  "conf.__constraints__.Char_Red", "0<=x<=255",
                  "conf.__constraints__.Char_Blue", "0<=x<=255",
                  "conf.__constraints__.Char_Green", "0<=x<=255",
                  "conf.__constraints__.Italic", "(0,1)",
                  "conf.__constraints__.Bold", "(0,1)",
                  "conf.__constraints__.Underline", "(0,1)",
                  "conf.__constraints__.Shadow", "(0,1)",
                  "conf.__constraints__.Strikeout", "(0,1)",
                  "conf.__constraints__.Contoured", "(0,1)",
                  "conf.__constraints__.Emphasis", "(0,1)",
                  "conf.__constraints__.Back_Red", "0<=x<=255",
                  "conf.__constraints__.Back_Blue", "0<=x<=255",
                  "conf.__constraints__.Back_Green", "0<=x<=255",
                  "conf.__constraints__.Code", "(utf-8,euc_jp,shift_jis)",
                  ""]




        
        



##
# @brief ユニコード文字列をドキュメント上で文字化けしない文字コードで文字列を返す
# @param m_str 変換前の文字列
# @return 変換後の文字列
#
def ResetCoding(m_str):
    if os.name == 'posix':
        return m_str.encode('utf-8')
    elif os.name == 'nt':
        return m_str.encode('cp932')








##
# @class OOoWriterControl
# @brief OpenOffice Writerを操作するためのRTCのクラス
#

class OOoWriterControl(WriterControl):
    ##
    # @brief コンストラクタ
    # @param self 
    # @param manager マネージャーオブジェクト
    #
  def __init__(self, manager):
    WriterControl.__init__(self, manager)
    
    

    try:
      self.writer = OOoWriter()
    except NotOOoWtiterException:
      return
    
    
    return


  


  ##
  # @brief 初期化処理用コールバック関数
  # @param self 
  # @return RTC::ReturnCode_t
  def onInitialize(self):
    WriterControl.onInitialize(self)
    
    OOoRTC.writer_comp = self

    
    return RTC.RTC_OK

  ##
  # @brief 文字書き込みの関数
  # @param self 
  # @param m_str 書き込む文字列
  #

  def setWord(self, m_str):
      cursor = self.writer.document.getCurrentController().getViewCursor()

      inp_str = OOoRTC.SetCoding(m_str, self.conf_Code[0])
      cursor.setString(inp_str)
      
      cursor.CharHeight = self.fontSize
      cursor.CharHeightAsian = self.fontSize

      

      if self.bold:
          cursor.CharWeight = BOLD
          cursor.CharWeightAsian = BOLD
      else:
          cursor.CharWeight = FW_NORMAL
          cursor.CharWeightAsian = FW_NORMAL
      if self.italic:
          cursor.CharPosture = ITALIC
          cursor.CharPostureAsian = ITALIC
      else:
          cursor.CharPosture = FS_NONE
          cursor.CharPostureAsian = FS_NONE

      if self.underline:
          cursor.CharUnderline = FU_SINGLE
      else:
          cursor.CharPosture = FU_NONE
          

      if self.strikeout:
          cursor.CharStrikeout = FST_SINGLE
      else:
          cursor.CharStrikeout = FST_NONE

      if self.emphasis:
          cursor.CharEmphasis = 1
      else:
          cursor.CharEmphasis = 0

      cursor.CharShadowed = self.shadow

      cursor.CharContoured = self.contoured
          

        

      #cursor.CharStyleName = self.fontName

      cursor.CharColor = OOoRTC.RGB(self.char_Red,self.char_Green,self.char_Blue)
      cursor.CharBackColor = OOoRTC.RGB(self.back_Red,self.back_Green,self.back_Blue)

      cursor.goRight(len(inp_str),False)

      cursor.collapseToEnd()

  ##
  # @brief カーソル位置の文字取得の関数
  # @param self
  # @return カーソル位置の文字列
  #

  def getWord(self):
      
      cursor = self.writer.document.getCurrentController().getViewCursor()

      try:
          out_str = ResetCoding(cursor.getString())
          return out_str
      except:
          return ""
       
      

  ##
  # @brief カーソルの位置を取得する関数
  # @param self 
  # @return カーソル位置のX座標、Y座標(単位はmm)
  #
  def oCurrentCursorPosition(self):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      oCurPos = cursor.getPosition()
      return oCurPos.X, oCurPos.Y

  ##
  # @brief カーソルをドキュメントの先頭に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStart(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoStart(sel)

  ##
  # @brief カーソルをドキュメントの最後尾に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEnd(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoEnd(sel)

  ##
  # @brief カーソルを行の先頭に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoStartOfLine(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoStartOfLine(sel)

  ##
  # @brief カーソルの行の最後尾に移動させる関数
  # @param self 
  # @param sel Trueならば移動範囲を選択
  #
  def gotoEndOfLine(self, sel):
      cursor = self.writer.document.getCurrentController().getViewCursor()
      cursor.gotoEndOfLine(sel)

  
  
  
      

  ##
  # @brief 文字数移動する関数
  # @param self 
  # @param diff 移動する文字数
  #
  def moveCharacter(self, diff):
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
  # @brief 単語数移動する関数
  # @param self 
  # @param diff 移動する単語数
  #
  def moveWord(self, diff):
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
  # @brief 行数移動する関数
  # @param self 
  # @param diff 移動する行数
  #
  def moveLine(self, diff):
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
  # @brief 段落数移動する関数
  # @param self 
  # @param diff 移動する段落数
  #
  def moveParagraph(self, diff):
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
  # @brief 活性化処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onActivated(self, ec_id):
    WriterControl.onActivated(self, ec_id)

    
    return RTC.RTC_OK

  def onDeactivated(self, ec_id):
    WriterControl.onDeactivated(self, ec_id)
    return RTC.RTC_OK

  
  ##
  # @brief 周期処理用コールバック関数
  # @param self 
  # @param ec_id target ExecutionContext Id
  # @return RTC::ReturnCode_t
  
  def onExecute(self, ec_id):
    WriterControl.onExecute(self, ec_id)
        

    return RTC.RTC_OK

  
  ##
  # @brief 終了処理用コールバック関数
  # @param self 
  # @param 
  # @return RTC::ReturnCode_t
  
  def onFinalize(self):
      WriterControl.onFinalize(self)
      OOoRTC.writer_comp = None
      return RTC.RTC_OK



##
# @brief コンポーネントを活性化してWriterの操作を開始する関数
#

def Start():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.mActivate()

##
# @brief コンポーネントを不活性化してWriterの操作を終了する関数
#

def Stop():
    
    if OOoRTC.writer_comp:
        OOoRTC.writer_comp.mDeactivate()


##
# @brief コンポーネントの実行周期を設定する関数
#

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
              
            OOoRTC.draw_comp.mSetRate(text)"""
      
      

      
        
        
      
      

      



      
  



##
# @brief RTCをマネージャに登録する関数
# @param manager マネージャーオブジェクト
def OOoWriterControlInit(manager):
  profile = OpenRTM_aist.Properties(defaults_str=ooowritercontrol_spec)
  manager.registerFactory(profile,
                          OOoWriterControl,
                          OpenRTM_aist.Delete)

##
# @brief
# @param manager マネージャーオブジェクト
def MyModuleInit(manager):
  manager._factory.unregisterObject(imp_id)
  OOoWriterControlInit(manager)

  
  comp = manager.createComponent(imp_id)






          

##
# @brief RTC起動の関数
#

def createOOoWriterComp():
    if OOoRTC.writer_comp:
        MyMsgBox('',OOoRTC.SetCoding('RTCは起動済みです','utf-8'))
        return                    
    
    if OOoRTC.mgr == None:
        if os.name == 'posix':
            home = expanduser("~")
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', home+'/OOoRTC/rtc.conf'])
        elif os.name == 'nt':
            OOoRTC.mgr = OpenRTM_aist.Manager.init([os.path.abspath(__file__), '-f', '.\\rtc.conf'])
        else:
            return

      
        OOoRTC.mgr.setModuleInitProc(MyModuleInit)
        OOoRTC.mgr.activateManager()
        OOoRTC.mgr.runManager(True)
    else:
        MyModuleInit(OOoRTC.mgr)
      
          

    try:
      writer = OOoWriter()
    except NotOOoWriterException:
      return

    
    MyMsgBox('',OOoRTC.SetCoding('RTCを起動しました','utf-8'))


    
    
    return None




##
# @brief メッセージボックス表示の関数
# @param title ウインドウのタイトル
# @param message 表示する文章
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

def MyMsgBox(title, message):
    try:
        m_bridge = Bridge()
    except:
        return
    m_bridge.run_infodialog(title, message)


##
# @brief OpenOfficeを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部
#

class Bridge(object):
  def __init__(self):
    self._desktop = XSCRIPTCONTEXT.getDesktop()
    self._document = XSCRIPTCONTEXT.getDocument()
    self._frame = self._desktop.CurrentFrame
    self._window = self._frame.ContainerWindow
    self._toolkit = self._window.Toolkit
  def run_infodialog(self, title='', message=''):
    try:
        msgbox = self._toolkit.createMessageBox(self._window,uno.createUnoStruct('com.sun.star.awt.Rectangle'),'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()
    except:
        msgbox = self._toolkit.createMessageBox(self._window,'infobox',1,title,message)
        msgbox.execute()
        msgbox.dispose()





##
# @brief OpenOffice Writerを操作するためのクラス
# http://d.hatena.ne.jp/kakurasan/20100408/p1のソースコード(GPLv2)の一部を改変
#

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
