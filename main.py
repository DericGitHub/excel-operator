import time
mc = 0
def pt(step):
    global mc 
    if mc == 0:
        mc = time.time()
    print 'step %s:takes %s'%(step,time.time()-mc)
    mc = time.time()
    pass

pt('load 1')
from controller.MainController import MainController,MainControllerUI
pt('load 2')
import multiprocessing
pt('load 3')
from multiprocessing import Queue,Process
pt('load 4')
def ui(queue_wr,queue_rd):
    pt('UI process 1')
#    from controller.MainController import MainControllerUI
    ui = MainControllerUI(queue_wr,queue_rd)
    pt('UI process 2')
    ui.run()
def controller(queue_wr,queue_rd):
    pt('CON process 1')
#    from controller.MainController import MainController
    controller = MainController(queue_wr,queue_rd)
    pt('CON process 1')
    controller.run()

def main():
    ui_queue = Queue()
    controller_queue = Queue()
    pui = Process(target=ui,args=(ui_queue,controller_queue))
    pcontroller = Process(target=controller,args=(controller_queue,ui_queue))
    pui.start()
    pcontroller.start()
    pt('main 3')
    while True:
        pui.join(timeout = 1)
        if not pui.is_alive():
            ui_queue.put(('stop',))
            pcontroller.join()
            break
     
if __name__ == '__main__':
    pt('main 1')
    multiprocessing.freeze_support()
    pt('main 2')
    app = main()

