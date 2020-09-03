import threading
class QOperator(object):

    def __init__(self):
        self.running = False
        self.workbook = None
        self.tasks = Fifo()
        self.threads = []
    # end __init__()

    def __worker_do__(self):
        pass

    def __on_worker_complete__(self):
        self.check_and_launch()
    # end __on_worker_complete__()
    
    def is_busy(self):
        for i in self.threads:
            if i.is_alive():
                return True
            # end if
        # end for
        return False
    # end is_busy

    def kill(self):
        self.running = False
    # end kill()

    def add_task(self, a_task):

        # put the task on the pile
        self.tasks.put(a_task)

        # if the system is not busy, 
        # run the tasks
        if not self.is_busy():
            self.running = True
            t = threading.Thread(target=self.__worker_do__)
            self.threads.append(t)

            t.start()
        # end if

    # end add_task

    def check_and_launch(self):
        if self.tasks.waiting() and self.running:
            t = threading.Thread(target=self.__worker_do__)
            self.threads.append(t)
            
            t.start()
        # end if
    # end check_and_launch
# end class


