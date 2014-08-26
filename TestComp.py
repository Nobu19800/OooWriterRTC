#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import OpenRTM_aist
import RTC
 
consolein_spec = ["implementation_id", "ConsoleIn",
                  "type_name",         "ConsoleIn",
                  "description",       "Console input component",
                  "version",           "1.0",
                  "vendor",            "sample",
                  "category",          "example",
                  "activity_type",     "DataFlowComponent",
                  "max_instance",      "10",
                  "language",          "Python",
                  "lang_type",         "script",
                  ""]
 
class ConsoleIn(OpenRTM_aist.DataFlowComponentBase):
    def __init__(self, manager):
        OpenRTM_aist.DataFlowComponentBase.__init__(self, manager)
        self._data = RTC.TimedString(RTC.Time(0,0),"")
        self._outport = OpenRTM_aist.OutPort("out", self._data)

        self.wordList = ["これは","OpenOffice Writer","を操作するRTCです"]
        self.count = 0
 
    def onInitialize(self):
        self.registerOutPort("out", self._outport)
        return RTC.RTC_OK
 
    def onExecute(self, ec_id):
        sys.stdin.readline()
        if self.count >= len(self.wordList):
            self.count = 0
        self._data.data = self.wordList[self.count]
        self.count += 1
        self._outport.write()
        
        return RTC.RTC_OK
 
def MyModuleInit(manager):
    profile = OpenRTM_aist.Properties(defaults_str=consolein_spec)
    manager.registerFactory(profile,
                            ConsoleIn,
                            OpenRTM_aist.Delete)
    comp = manager.createComponent("ConsoleIn")
 
def main():
    mgr = OpenRTM_aist.Manager.init(sys.argv)
    mgr.setModuleInitProc(MyModuleInit)
    mgr.activateManager()
    mgr.runManager()
 
if __name__ == "__main__":
    main()
