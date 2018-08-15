# -*- coding: utf-8 -*-

import sys
import os
import time
import struct
import pickle
import socket
from socketserver import ThreadingTCPServer, StreamRequestHandler
import logging

logging.basicConfig(filename=os.path.join(sys.path[0], 'EventLog.log'), level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s', filemode="w")
# logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s:%(message)s')


class TCP_Server(StreamRequestHandler):
    def send_data(self, data):
        data = pickle.dumps(data, True)
        self.request.sendall(struct.pack('>I', len(data)) + data)

    def recv_data(self):
        total_size = sys.maxsize
        sock_data = total_data = bytes()
        get_data_size = False
        while len(total_data) < total_size:
            sock_data = self.request.recv(1024)
            total_data += sock_data
            if not get_data_size and len(total_data) >= 4:
                total_size = struct.unpack('>I', total_data[:4])[0]
                total_data = total_data[4:]
                get_data_size = True
        return pickle.loads(total_data)

    def handle(self):
        logging.info('Client connected from %s', self.client_address)  # 输出客户端信息
        while True:
            try:
                fn = self.recv_data()
            except:
                logging.info('Connection is abnormally closed by %s', self.client_address)
                break

            if not fn:
                logging.info('Connection closed by %s', self.client_address)
                break

            try:
                with open(fn, 'rb') as f:
                    data = f.read()
                    logging.info('Reading the file of <%s> for sending.', fn)
            except:
                logging.warning('Opening file error:<%s>', fn)
                data = None

            try:
                self.send_data(data)
                logging.info('Send data to %s successfully.', self.client_address)
            except:
                logging.error('Sending data to %s failed.', self.client_address)
                break


def get_host_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
    finally:
        s.close()

    return ip


if __name__ == '__main__':
    time.sleep(30)
    try:
        host = get_host_ip()
        port = 10086
        addr = (host, port)
    except:
        logging.error('Failed to get local host name.')
        sys.exit(0)

    try:
        logging.info('Starting the server...')
        server = ThreadingTCPServer(addr, TCP_Server)
        logging.info('Server started, bind to %s', addr)
        server.serve_forever()
    except:
        logging.error('Server stopped!.')
