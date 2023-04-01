import threading
import qanew1
th = threading.Thread(target=qanew1.sending,args=['testdata.xlsx','qaquality456@gmail.com','qa1234567890','original34.jpg','4564','asdf','asdf'])
th.start()