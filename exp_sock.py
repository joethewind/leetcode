import socket
import threading

class tcp_server():
#creat socket server

    def __init__(self):
        self.host = '127.0.0.1'
        self.port = 8000
    def socketserver(self):
        server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        #bind host and port
        server.bind((self.host, self.port))
        #set detect time
        server.listen()
        # wait for client to connect
        # watch out：accept() will return a tuple
        # ele1 is client socket class，ele2 is client host and port
        clientsocket, addr = server.accept()

        #ensure a cycle
        while True:
            #receive client request
            recemsg = clientsocket.recv(1024)
            #decode data
            recdata = recemsg.decode('utf-8')
            #judge whether client send 'q',if yes,then quit
            if recdata == 'q':
                break
            print('receive:' + recdata)
            msg = input('reply:')
            clientsocket.send(msg.encode('utf-8'))







class tcp_client():
#creat socket client
    def __init__(self):
        self.host = '127.0.0.1'
        self.port = 8000

    def socketclient(self):
        #creat client class
        tcpclient = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        #connect
        tcpclient.connect((self.host, self.port))
        #ensure cycle
        while True:
            #input msg
            sendmsg = input('please input:')
            #if msg is 'q', then quit
            if sendmsg == 'q':
                break
            tcpclient.send(sendmsg.encode("utf-8"))
            msg = tcpclient.recv(1024)
            # receive server's reply
            print(msg.decode("utf-8"))


if __name__ == '__main__':
    server1 = tcp_server()
    client1 = tcp_client()
    threading.Thread(target=server1.socketserver).start()
    threading.Thread(target=client1.socketclient).start()
