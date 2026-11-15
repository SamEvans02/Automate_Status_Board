"""import socket


def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    print(ip)
    return ip
get_local_ip()"""










import openpyxl


wb = openpyxl.load_workbook("/mnt/c/projects/Automate_Status_Board/Team_Status_Board 2026.xlsx") #Load the existing Excel file 
sheet = wb.active #Access the active sheet 




def main():
    sheet["B7"] = "Test"
    
    return 0




wb.save("EDITED_Team_Status_Board 2026.xlsx") #Save the workbook 


if __name__ == "__main__":
    main()