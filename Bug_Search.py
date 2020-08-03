import requests
from oauth_token import get_token
import json
from excel_psirt import ciscopsirt
from time import sleep
import re
import openpyxl

disclaimer = '##############################################################################\n' \
             '#                                                                            #\n' \
             '#        Cisco IOS/IOSXE Bug Search tool developed by Sprint Networks        #\n' \
             '#        ************************************************************        #\n' \
             '#                                                                            #\n' \
             '#  This tool will help you find vulnerabilities in your Cisco IOS or IOSXE   #\n' \
             '#  system. The generated list of bugs will be exported to an Excel Sheet     #\n' \
             '#  for easy readability.                                                     #\n' \
             '#                             - Disclaimer -                                 #\n' \
             '#                                                                            #\n' \
             '#  This is only a trial version and will not store any of your information.  #\n' \
             '#  Sprint Networks do not take any responsibility of any information         #\n' \
             '#  provided using this tool. Any information provided will be sent directly  #\n' \
             '#  to Cisco Cloud to retrieve the relevant bugs from Cisco PSIRT.            #\n' \
             '#                                                                            #\n' \
             '#  As this is a trial version, you may not be able to retreive the bug       #\n' \
             '#  report if the daily limit has been exceeded. Contact Sprint Networks for  #\n' \
             '#  an indepth analysis of your network vulnerabilities.                      #\n' \
             '#                                                                            #\n' \
             '#  Version: 1.01V                                                            #\n' \
             '#  Developer: Sprint Networks - www.sprintnetworks.com                       #\n' \
             '#  Contact: info@sprintnetworks.com                                          #\n' \
             '#                                                                            #\n' \
             '##############################################################################\n'

def excel_gen(version, token, platform='IOS'):
    # 1 for IOS && 2 for IOSXE
    print('Generating report for ' + version)
    # token = get_token()
    if platform == 'IOS':
        http_request = "https://api.cisco.com/security/advisories/ios?version=" + version
    elif platform == 'IOSXE':
        http_request = "https://api.cisco.com/security/advisories/iosxe?version=" + version
    else:
        print('Unsupport Platform')
    r = requests.get(
        http_request,
        headers={'Authorization': token}
        )
    result = json.loads(r.content)
    if r.status_code == 200:
        ciscopsirt(result, name=version)
        print('Report for ' + version + ' has been done')
        sleep(2)
    else:
        print(r.status_code)
        print('Wrong Input\n'
              'Please restart this program')
        sleep(2)
    print('\nPlease find the results under the same folder!\n')
    sleep(3)

def platform_get():
    platform = ''
    while platform != '1' or platform != '2':
        platform = input(
                         'Enter the number to choose software platform:\n'
                         '1: IOS \n'
                         '2: IOSXE\n'
                         ': '
        )
        if platform == '1':
            return 'IOS'
        elif platform == '2':
            return 'IOSXE'
        else:
            print('Wrong Input\n')


def version_list_get(platform='IOS'):
    version_str = input('Enter the list of Cisco ' + platform + ' you want to check, seperated with comma, '
                                                                'e.g.,12.2(50)SE1,15.2(4)M5\n')
    version_list = re.split(',', version_str)

    return version_list

def main():
    print("System is initialising")

    token = get_token()
    print(disclaimer)
    sleep(3.5)

    mode = input('Enter the number to choose the mode,: \n'
                 '1: Input one Cisco software version\n'
                 '2: Input multiple Cisco software versions\n'
                 '3: Exit\n'
                 ':')
    if mode == '2':
        platform = platform_get()
        version_list = version_list_get()
        if platform == 'IOS':
            for ios_version in version_list:
                # print('Generating report for ' + ios_version)
                excel_gen(ios_version, token, platform=platform)
        elif platform == 'IOSXE':
            for iosxe_version in version_list:
                # print('Generating report for ' + iosxe_version)
                excel_gen(iosxe_version, token, platform=platform)
        else:
            print('No valid version has been found')

    elif mode == '1':
         platform = platform_get()
         version = input('Enter the software version of your Cisco device, e.g.,12.2(50)SE1 :')
         excel_gen(version, token, platform=platform)


    elif mode == '3':
        exit()
    else:
        print('Wrong Input')
    # print('\n')

    print('Program is exiting')
    sleep(3)

if __name__ == '__main__':
    main()