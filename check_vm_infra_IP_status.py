#!/usr/bin/env python3

import subprocess
import sys
import openpyxl

args = sys.argv

if len(sys.argv) == 1:
    print("Please provide the Start IP range and End IP range.")
    print("")
    print("     Ex: python3 host_alive 192.168.10.1 192.168.100")
elif len(sys.argv) > 5:
    print("Too many arguments.")
    print("")
    print("     Ex: python3 host_alive 192.168.10.1 192.168.100")

else:
    Start_IP = args[1]
    End_IP = args[2]

    Start_IP_Last_Octect = int(Start_IP.split(".")[3])
    End_IP_Last_Octect = End_IP.split(".")[3]
    
    IP_Range = int(End_IP_Last_Octect) - int(Start_IP_Last_Octect) + 1
    
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    worksheet["A1"] = "IP Address"
    worksheet["B1"] = "Status"

    column_value = 2

    for i in range(IP_Range):
        
        IP = Start_IP.split(".")[0] + "." + Start_IP.split(".")[1] + "." +Start_IP.split(".")[2] + "." + str(Start_IP_Last_Octect)
        
        PING = "ping -c 1 " + IP
        output = subprocess.getoutput(PING)

        IP_Place_value = "A" + str(column_value)
        Status_Place_value = "B" + str(column_value)

        worksheet[IP_Place_value] = IP
        worksheet[Status_Place_value]
        
        if "From 198.18.19.170 icmp_seq=1 Destination Host Unreachable" in output:
            print(IP + " - " + " Dropping from the VM Gateway - 198.18.19.170")
            worksheet[Status_Place_value] = "Host Unreachable"
        elif "Request timed out." in output:
            print(IP + " - " + " Request Timed Out")
            worksheet[Status_Place_value] = "Host Timed Out"
        else:
            print(IP + " - " + " Host is Up")
            worksheet[Status_Place_value] = "Host is Up"
        
        Start_IP_Last_Octect += 1
        column_value += 1
    
    if len(sys.argv) == 4:
        if args[3] == "--save":

            doc_name = str(Start_IP) + "-" + str(End_IP) + "-IP-Status.xlsx"
            print("\nIP Status was written to " + doc_name + ".")
            workbook.save(doc_name)

