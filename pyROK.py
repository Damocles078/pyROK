import os,sys
sys.path.append(os.path.dirname(__file__))
from lib.ROK_GUI import pyROK_GUI
from lib.ROK_analyser import ROKanalyzer
from lib.ROK_puller import ROKpuller
import threading
import time
from elevate import elevate

def run():
    pyROK = pyROK_GUI()
    GUI = threading.Thread(target=pyROK.pyROK_run)
    GUI.start()

    while not pyROK.close.is_set() and not pyROK.configDone.is_set():
        time.sleep(1)

    if pyROK.configDone.is_set():

        (write_request, table_name) = pyROK.generate_sql_commands()

        analyzer = ROKanalyzer(request=write_request, 
                            GUI=pyROK,
                            acquire_alliance=pyROK.acquire_alliance,
                            acquire_dead=pyROK.acquire_dead,
                            acquire_name=pyROK.acquire_playerNane,
                            acquire_power=pyROK.acquire_power,
                            acquire_rssAssist=pyROK.acquire_rssAssist,
                            acquire_T4=pyROK.acquire_T4,
                            acquire_T5=pyROK.acquire_T5)
        
        puller = ROKpuller(GUI=pyROK,analyzer=analyzer)

        analyzer_thread = threading.Thread(target=analyzer.worker)
        analyzer_thread.start()
        puller.Acquire()

def main():
    elevate(show_console=False)
    run()

if __name__=='__main__':
    main()

