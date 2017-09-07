from controller import MainController
from multiprocessing import Queue,Process

def ui(queue_wr,queue_rd):
    ui = MainController.MainControllerUI(queue_wr,queue_rd)
    ui.run()
def controller(queue_wr,queue_rd):
    controller = MainController.MainController(queue_wr,queue_rd)
    controller.run()

def main():
    ui_queue = Queue()
    controller_queue = Queue()
    pui = Process(target=ui,args=(ui_queue,controller_queue))
    pcontroller = Process(target=controller,args=(controller_queue,ui_queue))
    pui.start()
    pcontroller.start()
    while True:
        pui.join(timeout = 1)
        if not pui.is_alive():
            ui_queue.put(('stop',))
            pcontroller.join()
            break
    

if __name__ == '__main__':
    app = main()

