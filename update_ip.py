# -*- coding: utf-8 -*-
#!/usr/bin/env python3
import os
import re
import time
import socket

##
# 定期每30秒解析一下域名，将IP写入proxychains.conf文件，达到动态域名解析效果，配合proxychains使用的
# #
def get_ip(domain):
    try:
        ip = socket.gethostbyname(domain)
        return ip
    except socket.gaierror:
        return None

def replace_ip_in_file(file_path, old_ip, new_ip):
    with open(file_path, 'r') as file:
        content = file.read()
    content = re.sub(r'\b' + old_ip + r'\b', new_ip, content)
    with open(file_path, 'w') as file:
        file.write(content)

def main():
    domain = 'www.19940701.xyz'
    file_path = '/etc/proxychains.conf'
    regex = r'socks5\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+20808'

    while True:
        new_ip = get_ip(domain)
        if new_ip:
            with open(file_path, 'r') as file:
                content = file.read()
            match = re.search(regex, content)
            if match:
                old_ip = match.group(1)
                if old_ip != new_ip:
                    replace_ip_in_file(file_path, old_ip, new_ip)
                    print(f'IP 地址已更新: {old_ip} -> {new_ip}')
                else:
                    print(f'IP 地址未变化: {old_ip}')
            else:
                print(f'未找到匹配的 IP 地址')
        else:
            print(f'无法解析域名: {domain}')
        
        time.sleep(30)

if __name__ == '__main__':
    main()
