import xlwings as xw
import numpy as np
import pandas as pd 
import os
import yaml
import logging
import colored
import getpass


def get_config() -> dict:
    
    try:  
        
        with open(get_Project_Directory() + r'/config/configuration.yaml') as stream:
            config = yaml.safe_load(stream)
    except yaml.YAMLError as e:
        print(colored.fg('red'))
        print("Failed to parse configuration yaml " + e.__str__())
        print(colored.fg('white'))
        logging.error("Failed to parse configuration yaml " + e.__str__())
                
    except Exception as e:
        print(colored.fg('red'))
        print("Failed to parse configuration yaml " + e.__str__())
        print(colored.fg('white'))
        logging.error("Failed to parse configuration yaml " + e.__str__())
    
    return config
    
def get_settings() -> dict:
    
    try:        
        with open(get_Project_Directory() + r'/config/settings.yaml') as stream:
            settings = yaml.safe_load(stream)
    except yaml.YAMLError as e:
        print(colored.fg('red'))
        print("Failed to parse settings yaml " + e.__str__())
        print(colored.fg('white'))
        logging.error("Failed to parse settings yaml " + e.__str__())
                
    except Exception as e:
        print(colored.fg('red'))
        print("Failed to parse settings yaml " + e.__str__())
        print(colored.fg('white'))
        logging.error("Failed to parse settings yaml " + e.__str__())
    
    return settings

def get_Project_Directory():
    return os.path.dirname(os.path.realpath('CATC.py'))
    
def get_value(creds : dict, pathtype : str):    
    return creds[pathtype.lower()]['value']

def set_Log_Directory():
    file_Path = get_Project_Directory() + '/Log'

    if os.path.exists(file_Path) == False:
        os.mkdir(file_Path)
    
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    logging.basicConfig(filename=file_Path + '/CATCLog.log',
                        filemode='w',
                        format='%(asctime)s|%(lineno)d|%(levelname)s|%(message)s',
                        datefmt='%Y_%m_%d_%H_%M_%S',
                        level=logging.DEBUG)
    
def get_market(geo:str)->str:
    
    geo_dict = {
        'AM' : ['US','C','LA'],
        'AP' : ['ANZ','ASE','CHR','ISA','KOR'],
        'JN' : ['JN'],
        'EM' : ['BNL','CEE','DAC','FRA','IGI','ITY','NFR','UKI','MER']        
    }     
    return geo_dict.get(geo,"")

def get_GEO()->str:
    
    ww_dict = {
        'WW' : ['AM','AP','JN','EM'] 
    }
    return ww_dict.get('WW',"")

def get_WW(ww:str='WW')->str:
    
    return 'WW'

def get_user_details(userid:str)->str:
    
    user_dict = {        
        'WWUSER' : ['WWPWD', 'WW'],
        'AMUSER' : ['AMPWD', 'AM'],
        'EMUSER' : ['EMPWD', 'EM'],
        'APUSER' : ['APPWD', 'AP'],
        'JNUSER' : ['JNPWD', 'JN'],
    }
    
    return user_dict.get(userid.upper(),"Invalid".capitalize())

def validate_market_file_path(market:dict)->bool:
    
    validate = 1
    
    try :
        config = get_config()
        
        for key in market:
            
            file_path = get_value(config,  "path_market_" + key.lower() )
            print("\nFile Path from Config file for " +  key +": " + file_path)
            logging.info(file_path)       
            
            if validate_file_path(file_path) == False:
                print(colored.fg('red'))
                print("\nFile Path Doesn't exits : " + file_path)
                print(colored.fg('white'))
                logging.error(file_path)
                validate+=1         
    except Exception as e:
        validate+=1
        print(colored.fg('red'))
        print("Exception : " + e.__str__())
        print(colored.fg('white'))
        logging.error("Exception : " + e.__str__())
        
    return True if validate  == 1 else False
    

def validate_geo_file_path(geo:dict)->bool:
    
    validate = 1
    
    try:
        config = get_config()
        
        for key in geo:
            
            file_path = get_value(config,  "path_geo_" + key.lower() )
            print("\nFile Path from Config file for " +  key +": " + file_path)
            logging.info(file_path)       
            
            if validate_file_path(file_path) == False:
                print(colored.fg('red'))
                print("\nFile Path Doesn't exits : " + file_path)
                print(colored.fg('white'))
                logging.error(file_path)
                validate+=1         
       
    except Exception as e:
        validate+=1
        print(colored.fg('red'))
        print("Exception : " + e.__str__())
        print(colored.fg('white'))
        logging.error("Exception : " + e.__str__())
        
    return True if validate  == 1 else False
    
def validate_ww_file_path(ww:str)->bool:
    
    validate = 1
    try:
        config = get_config()
        
        for key in ww:
            
            file_path = get_value(config,  "path_" + key.lower() )
            print("\nFile Path from Config file for " +  key +": " + file_path)
            logging.info(file_path)       
            
            if validate_file_path(file_path) == False:
                print(colored.fg('red'))
                print("\nFile Path Doesn't exits : " + file_path)
                print(colored.fg('white'))
                logging.error(file_path)
                validate+=1       
                
    except Exception as e:
        validate+=1
        print(colored.fg('red'))
        print("Exception : " + e.__str__())
        print(colored.fg('white'))
        logging.error("Exception : " + e.__str__()) 
        
    return True if validate  == 1 else False
       
def validate_file_path(file_path:str)->bool:
    return os.path.exists(file_path)


def login()->[str,bool]:

    logging.info("Login")
    
    login_attempt = 0
    while login_attempt < 3 :
        
        username=input("\n **Please enter your User Id: ")    
        
        logging.info("Entered username : " + username)
        
        if get_user_details(username) != 'Invalid'.capitalize():
            print(colored.fg('green'))
            print("--- Welcome " + username + " ---")
            print(colored.fg('white'))
            attempts = 0
            logging.info("login attempt : " + attempts.__str__())
            
            while attempts<3:

                print("\n **Please enter your password")
                password  = getpass.getpass(prompt="Password :")
                if password.upper() == get_user_details(username)[0].upper():
                    print(colored.fg('green'))
                    print("\n You have successfully completed Log in")
                    print("--Hello",username)
                    print(colored.fg('white'))
                    print('--Executive File starts runninng')
                    logging.info("You have successfully completed Log in")
                    logging.info("--Hello "+ username)
                    logging.info("Executive File starts running")
                    login_result = 1
                    return [str(get_user_details(username)[1]), True]            
                    break
                else:
                    print(colored.fg('red'))
                    print("\n Incorrect password. Please enter your password again")    
                    attempts+=1
                    print(colored.fg('yellow'))                    
                    print("--Attempts left: ",3-attempts)
                    print(colored.fg('white'))
                    logging.info("Incorrect password. Please enter your password again")
                    logging.info("login attempt : " + attempts.__str__())                 
                    if attempts>=3:
                        print(colored.fg('red'))
                        print("\n You have attemplted too many")
                        print("--You can reach me via e-mail/slack 'jinmun.park@ibm.com' to have any questions using the application.")
                        print(colored.fg('white'))
                        logging.info("You have made too many attempts")
                        login_result = 0
                        logoff()
                        return ["Invalid User", False]            
        else:
            print(colored.fg('red'))
            print("\n Your User Id is Unknown User. Please try again")
            print(colored.fg('white'))
            logging.info("Your User Id is Unknown User. Please try again")
            login_result = 0     
            login_attempt +=1
            if( login_attempt >= 3):
                print(colored.fg('red'))
                print("\n You have attemplted too many")
                print("--You can reach me via e-mail/slack 'jinmun.park@ibm.com' to have any questions using the application.")
                print(colored.fg('white'))
                logging.info("You have made too many attempts")
                login_result = 0
                logoff()
                return ["Invalid User", False] 
 
def logoff():
    print("Log Out")
    logging.info("Logging off")    

    

def run_market_to_GEO(key_input : str):
    
    print(colored.fg('green'))
    print("Progressing Market to GEO")
    print(colored.fg('white'))
    
    #### SETUP 001 : Market to Geo Dictionary
    
    market_dc = {'AM' : ['US', 'LA', 'C'],
                'EM' : ['BNL','CEE','DAC','FRA','IGI','ITY','NFR','UKI','MER'], 
                'JN' : ['JN'],
                'AP' : ['ANZ','ASE','CHR','ISA','KOR']
                }
    
    worksheet_dc = {'name' : ['AM', 'EM', 'JN', 'AP', 'US', 'LA', 'C','BNL','CEE','DAC','FRA','IGI','ITY','NFR','UKI','MER','JN','ANZ','ASE','CHR','ISA','KOR'],
                    'main' : [None,None,None,None, 'US', 'LA', 'C','BNL','CEE','DAC','FRA','IGI','ITY','NFR','UKI','MER','JN','ANZ','ASE','CHR','ISA','KOR'],
                    'bridge' : [None,None,None,None, 'US_Bridge', 'LA_Bridge', 'C_Bridge','BNL_Bridge','CEE_Bridge','DAC_Bridge','FRA_Bridge','IGI_Bridge','ITY_Bridge','NFR_Bridge','UKI_Bridge','MER_Bridge','JN_Bridge','ANZ_Bridge','ASE_Bridge','CHR_Bridge','ISA_Bridge','KOR_Bridge'],
                    'roadmap' : [None,None,None,None, 'US_Roadmap', 'LA_Roadmap', 'C_Roadmap','BNL_Roadmap','CEE_Roadmap','DAC_Roadmap','FRA_Roadmap','IGI_Roadmap','ITY_Roadmap','NFR_Roadmap','UKI_Roadmap','MER_Roadmap','JN_Roadmap','ANZ_Roadmap','ASE_Roadmap','CHR_Roadmap','ISA_Roadmap','KOR_Roadmap'],
                    'togo' : [None,None,None,None, 'US_To_Go', 'LA_To_Go', 'C_To_Go','BNL_To_Go','CEE_To_Go','DAC_To_Go','FRA_To_Go','IGI_To_Go','ITY_To_Go','NFR_To_Go','UKI_To_Go','MER_To_Go','JN_To_Go','ANZ_To_Go','ASE_To_Go','CHR_To_Go','ISA_To_Go','KOR_To_Go'],
                    'signing' : [None,None,None,None, 'US_Signing', 'LA_Signing', 'C_Signing','BNL_Signing','CEE_Signing','DAC_Signing','FRA_Signing','IGI_Signing','ITY_Signing','NFR_Signing','UKI_Signing','MER_Signing','JN_Signing','ANZ_Signing','ASE_Signing','CHR_Signing','ISA_Signing','KOR_Signing']
                    }
    
    worksheet_df = pd.DataFrame(worksheet_dc, columns = ['name','main','bridge','roadmap','togo','signing'])
    worksheet_df.set_index("name", inplace = True)
    
    
    #### SETUP 002 : User Iput
    
    print('\n The fie has a list of Geo in Main Scripts ; ')
    logging.info("The fie has a list of Geo in Main Scripts")
    for key, value in market_dc.items() :
        print(key)
    
    print("--Your user setting will select" + " [" + str(key_input) + "].")
    
    logging.info("Your user setting will select : " + str(key_input))
    
    select_dc = dict(map(lambda key: (key, market_dc.get(key, None)), [key_input]))
    
    print(select_dc)
    for key, value in select_dc.items() :
        print('\n Selected Geo have following markets in the list. Please check your Local Direcotry YAML.file to run the main script accordingly')
        print(value)
        logging.info("Selected Geo have following markets in the list. Please check your Local Direcotry YAML.file to run the main script accordingly")
        logging.info(value)
    
    #### SETUP 003 : Filterign Dictionary base on the user selection
    
    for key, value in select_dc.items() :
        select_geo = key
        select_market = list(value)
    
    ws_market_list = []
    for i in range(len(select_market)) :
        ws_market = list(worksheet_df.loc[str(select_market[i])])
        ws_market_list.append(ws_market[0:len(ws_market)])
    
    ws_market_df = pd.DataFrame(ws_market_list, columns = ['main','bridge','roadmap','togo','signing'])     
        
    #### MAIN : COPY & PASTE STARTS
         
    try:        
        
        app = xw.App(visible=False, add_book=False)    
        
        config = get_config()
        settings = get_settings()
        
        #### MAIN 001 : Range_Setup
            
        print("\nRunning Main Script : Setting File")
        logging.info("Running Main Script : Setting File")
            
        main_setup = get_value(settings,'main_setup')
        bridge_setup = get_value(settings,'bridge_setup')
            
        roadmap_setup_a = get_value(settings,'roadmap_setup_a')
        roadmap_setup_b = get_value(settings,'roadmap_setup_b')
        
        togo_setup = get_value(settings,'togo_setup')
        
        signing_setup = get_value(settings,'signing_setup') 
        
        print("\nCompleted running Setting File")
        logging.info("Completed running Setting File")
   
        #### MAIN 002 : Call Geo/Market Path and Load Geo/Market_Worksheets
            
        # Call Geo Path
        print("\nRunning Main Script : loadinng geo worksheets")
        logging.info("Running Main Script : loadinng geo worksheets")
        
        path_geo = get_value(config, str('path_geo_' + select_geo.lower()))
            
        # Load Geo Workbook
        geo_wb = app.books.open(path_geo)
        
        #### MAIN 003 : Call Market Path + Load Geo/Markeht Woksheets    
        print("\nRunning Main Script : loadinng market worksheets")
        logging.info("Running Main Script : loadinng market worksheets")
        
        call_wb = {}            
        
        for i in range(len(select_market)) : 
            
            print("\nOpening Excel " +  'market_wb_{0}'.format(i) + " in Background")
            
            # Callig Path_Market
            call_wb[i] = get_value(config, str('path_market_' + select_market[i].lower()))
                
            # Load Market Workbook
            globals()['market_wb_{0}'.format(i)] = app.books.open(call_wb[i])
            
            wb = globals()['market_wb_{0}'.format(i)]
            wb_name = str('market_wb_{0}'.format(i))
                        
            # Calling Worksheet Names in Market Templates
            worksheet_market = ws_market_df.iloc[i]
                
            # Calling Worksheet Names in Geo Template
            worksheet_geo = ws_market_df.iloc[i]
                            
            # Loading Geo/Market Worksheets
            for j in range(len(ws_market_df.columns)) :  
                 
                # market_wb_{0}_ws_{j} 
                globals()[str(wb_name) + '_ws_{0}'.format(j)] = wb.sheets[worksheet_market[j]]
                print("\nProcessing worksheet " + 'market_wb_{0}'.format(i) + '_ws_{0}'.format(j))
                
                # geo_mar_{0}_ws_{j}  
                globals()['geo_mar_{0}'.format(i) + '_ws_{0}'.format(j)] = geo_wb.sheets[worksheet_geo[j]]
                print("\nProcessing worksheet " + 'geo_mar_{0}'.format(i) + '_ws_{0}'.format(j))
        
        print("\nCompleted loading market worksheets")
        logging.info("Completed loading market worksheets")
        
        #### MAIN 004 : Copy and Paste (Market -> Geo)
            
        ###############################################################################################################
        
             ## INDEX FOR (Market -> Geo)
                # geo_mar_{0}_ws_{j}   : {j = 0} =  main  / {j = 1} = bridge  / {j = 2} = roadmap / {j = 3} = togo / {j = 4} = signing 
                # market_wb_{0}_ws_{j} : {j = 0} =  main  / {j = 1} = bridge  / {j = 2} = roadmap / {j = 3} = togo / {j = 4} = signing
        
        ###############################################################################################################
            
        print("\nRunning Main Script : Working copy and paste from Market to Geo")
        
        logging.info("Running Main Script : Working copy and paste from Market to Geo")
        
        def range_call(ws_number : int):
            if ws_number == int(0):
                return main_setup
            elif ws_number == int(1):
                return bridge_setup
            elif ws_number == int(2):
                return roadmap_setup_a
            elif ws_number == int(3):
                return togo_setup
            else:
                return signing_setup
        
        # loop copy (roadmap range_setup_a)
        for i in range(len(select_market)) : 
            for j in range(len(ws_market_df.columns)) : 
            # Selction from {j = 0, main} to {j = 4, signing_setup}
              
                # Loop the rane of cells to be copied
                range_load = range_call(ws_number = j)
                print("\nCopying from " + 'market_wb_{0}'.format(i) + '_ws_{0}'.format(j) + " to " + 'geo_mar_{0}'.format(i) + '_ws_{0}'.format(j))
                logging.info("Copying from " + 'market_wb_{0}'.format(i) + '_ws_{0}'.format(j) + " to " + 'geo_mar_{0}'.format(i) + '_ws_{0}'.format(j))
                
                # Copy the range selected from market sheets
                copy = globals()['market_wb_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).options(numbers=int).value
                
                # Paste to Geo worksheets
                globals()['geo_mar_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).value = copy
                
                # Center Allignment
                globals()['geo_mar_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).api.HorizontalAlignment = -4108    
        
        # manual copy (roadmap range_setup_b)        
        for i in range(len(select_market)) : 
            
            copy_manual = globals()['market_wb_{0}'.format(i) + '_ws_2'].range(roadmap_setup_b).options(numbers=int).value
            globals()['geo_mar_{0}'.format(i) + '_ws_2'].range(roadmap_setup_b).value = copy_manual
            globals()['geo_mar_{0}'.format(i) + '_ws_2'].range(range_load).api.HorizontalAlignment = -4108    
            print("\nCopying from " + 'market_wb_{0}'.format(i) + '_ws_2' + " to " + 'geo_mar_{0}'.format(i) + '_ws_2')
            logging.info("Copying from " + 'market_wb_{0}'.format(i) + '_ws_2' + " to " + 'geo_mar_{0}'.format(i) + '_ws_2')
        
        print(colored.fg('green'))
        print("\nCompleted copying worksheets from Market to Geo")
        print("\nCompleted running Main Script")
        print(colored.fg('white'))
        logging.info("Completed copying worksheets from Market to Geo")
        logging.info("Completed running Main Script")
        
        #### MAIN 005 : Save and Close
         
        print("\n\nSaving all worksheets and close the workbook")
        logging.info("Saving all worksheets and close the workbook")
        
        # Set Directory Path for the first market template and save
        globals()['market_wb_0'].save(call_wb[0])
        print("Saving Workbook " + 'market_wb_0' )
        globals()['market_wb_0'].close()
        print("Closing Workbook " + 'market_wb_0')
        
        # Save the rest market templates to the path 
        for i in range(1, len(select_market)) : 
            print("Saving Workbook " + 'market_wb_{0}'.format(i))
            logging.info("Saving Workbook " + 'market_wb_{0}'.format(i))
            globals()['market_wb_{0}'.format(i)].save()
            
            print("Closing Workbook " + 'market_wb_{0}'.format(i) )
            logging.info("Closing Workbook " + 'market_wb_{0}'.format(i))
            globals()['market_wb_{0}'.format(i)].close()
        
        # Save Geo Template
        geo_wb.save(path_geo)
        geo_wb.save()
        geo_wb.close() 
        
        app.quit()
        
        print(colored.fg('green'))
        print("Completed running all scripts. Please check the workbook")
        print(colored.fg('white'))
        logging.info("Completed running all scripts. Please check the workbook")
        
        
    except Exception as e:
        print(colored.fg('red'))
        print("Exception : " + e.__str__())
        print(colored.fg('white'))
        logging.error("Exception : " + e.__str__())
        app.quit()   

def run_GEO_to_Worldwide(key_input : str):
    
    print(colored.fg('green'))
    print("Progressing Geo to WorldWide Template")
    print(colored.fg('white'))
    logging.info("Progressing Geo to WorldWide Template")
    
    #### SETUP 001 : Geo to WW Dictionary
    worldwide_dc = {'WW' : ['AM', 'EM', 'JN', 'AP']}
    
    worksheet_dc = {'name' : ['WW', 'AM', 'EM', 'JN', 'AP'],
                   'main' : ['WW', 'AM', 'EM', 'JN', 'AP'],
                   'bridge' : ['WW_Bridge', 'AM_Bridge', 'EM_Bridge', 'JN_Bridge', 'AP_Bridge'],
                   'roadmap' : ['WW_Roadmap', 'AM_Roadmap', 'EM_Roadmap', 'JN_Roadmap', 'AP_Roadmap'],
                   'revenue' : ['WW_Rev', 'AM_Rev', 'EM_Rev', 'JN_Rev', 'AP_Rev'],
                   'togo' : ['WW_To_Go', 'AM_To_Go', 'EM_To_Go', 'JN_To_Go', 'AP_To_Go'],
                   'signing' : ['WW_Signing', 'AM_Signing', 'EM_Signing', 'JN_Signing', 'AP_Signing']
                   }
    
    worksheet_df = pd.DataFrame(worksheet_dc, columns = ['name','main','bridge','roadmap','revenue','togo','signing'])
    worksheet_df.set_index("name", inplace = True)
    
    #### SETUP 002 : User Iput
    
    print('\n The fie has a list of Geo in Main Scripts ; ')
    logging.info("The fie has a list of Geo in Main Scripts")
   
    for key, value in worldwide_dc.items() :
        print(key)
    
    print("--Your user setting will select" + " [" + str(key_input) + "].")
    logging.info("Your user setting will select : " + str(key_input))
    
    select_dc = dict(map(lambda key: (key, worldwide_dc.get(key, None)), [key_input]))
    
    print(select_dc)
    for key, value in select_dc.items() :
        print('\n Selected Geo have following markets in the list. Please check your Local Direcotry YAML.file to run the main script accordingly')
        print(value)
        logging.info("Selected Geo have following markets in the list. Please check your Local Direcotry YAML.file to run the main script accordingly")
        logging.info(value)
    
    #### SETUP 003 : Filterign Dictionary base on the user selection
    
    for key, value in select_dc.items() :
        select_geo = key
        select_market = list(value)
    
    ws_geo_list = []
    ws_geo = list(worksheet_df.loc[str(select_geo)])
    ws_geo_list.append(ws_geo[0:len(ws_geo)])
    ws_geo_df = pd.DataFrame(ws_geo_list, columns = ['main','bridge','roadmap','revenue', 'togo', 'signing'])
    
    ws_market_list = []
    for i in range(len(select_market)) :
        ws_market = list(worksheet_df.loc[str(select_market[i])])
        ws_market_list.append(ws_market[0:len(ws_market)])
    
    ws_market_df = pd.DataFrame(ws_market_list, columns = ['main','bridge','roadmap','revenue','togo','signing'])
      
        
    #### MAIN : COPY & PASTE STARTS
    
    try:        
        
        app = xw.App(visible=False, add_book=False)    
        
        config = get_config()
        settings = get_settings()
        
        
        #### MAIN 001 : Range_Setup
            
        print("\nRunning Main Script : Setting File")
        logging.info("Running Main Script : Setting File")
        
        main_setup = get_value(settings,'main_setup')
        bridge_setup = get_value(settings,'bridge_setup')
            
        roadmap_setup_a = get_value(settings,'roadmap_setup_a')
        roadmap_setup_b = get_value(settings,'roadmap_setup_b')
        
        revenue_setup = get_value(settings,'revenue_setup')
        
        togo_setup = get_value(settings,'togo_setup')
        
        signing_setup = get_value(settings,'signing_setup') 
        
        print("\nCompleted running Setting File")
        logging.info("Completed running Setting File")
        

        #### MAIN 002 :  Call Market Path + Load Geo/Markeht Woksheets    
            
        print("\nRunning Main Script : loadinng WW/Geo worksheets")
        logging.info("Running Main Script : loadinng WW/Geo worksheets")
        
        call_wb = {}            
        
        for i in range(len(select_market)) : 
            
            # Calling Path_WW
            path_geo = get_value(config, str('path_' + select_geo.lower()))
            
            # Calling Path_Geo
            call_wb[i] = get_value(config, str('path_geo_' + select_market[i].lower()))
                
            # Load WW Workbook
            geo_wb = app.books.open(path_geo)
            
            # Load Geo Workbook
            globals()['geo_wb_{0}'.format(i)] = app.books.open(call_wb[i])
            print("\nOpening Excel " +  'geo_wb_{0}'.format(i) + "in Background")
            wb = globals()['geo_wb_{0}'.format(i)]
            wb_name = str('geo_wb_{0}'.format(i))
                        
            # Load WW/GEO Worksheet Names
            worksheet_market = ws_market_df.iloc[i] #Geo Worksheets
            
            # Load WW/GEO Worksheets in Python
            for j in range(len(ws_market_df.columns)) :     
                 
                # PART 1 : Loop Geo
                globals()[str(wb_name) + '_ws_{0}'.format(j)] = wb.sheets[worksheet_market[j]]
                
                print("\nProcessing Geo worksheet " + 'geo_wb_{0}'.format(i) + '_ws_{0}'.format(j))
                
                # PART 2 : Loop Worldwide
                globals()['ww_geo_{0}'.format(i) + '_ws_{0}'.format(j)] = geo_wb.sheets[worksheet_market[j]]
                
                print("\nProcessing WW worksheet " + 'ww_geo_{0}'.format(i) + '_ws_{0}'.format(j))
               

        print("\nCompleted loading WW/Geo worksheets")
        logging.info("Completed loading WW/Geo worksheets")
    
                     
        #### MAIN 004 : Copy and Paste (Market -> Geo)
        print("\nRunning Main Script : Working copy and paste from Geo to WW")
        logging.info("Running Main Script : Working copy and paste from Geo to WW")
        
        def range_call(ws_number : int):
            if ws_number == int(0):
                return main_setup
            elif ws_number == int(1):
                return bridge_setup
            elif ws_number == int(2):
                return roadmap_setup_a
            elif ws_number == int(3):
                return revenue_setup
            elif ws_number == int(4):
                return togo_setup
            else:
                return signing_setup     
        
            range_call(ws_number=5)

        # loop copy (roadmap range_setup_a)
        for i in range(len(select_market)) : 
            for j in range(len(ws_market_df.columns)) :  
            # Selction from {j = 0} =  main  / {j = 1} = bridge  / {j = 2} = roadmap / {j = 3} = revenue / {j = 4} = togo / {j = 5} = signing
              
                # Loop the rane of cells to be copied
                range_load = range_call(ws_number = j)
                    
                print("\nCopying from " + 'geo_wb_{0}'.format(i) + '_ws_{0}'.format(j) + " to " + 'ww_geo_{0}'.format(i) + '_ws_{0}'.format(j))
                logging.info("Copying from " + 'geo_wb_{0}'.format(i) + '_ws_{0}'.format(j) + " to " + 'ww_geo_{0}'.format(i) + '_ws_{0}'.format(j))
                
                # Copy the range selected from Geo sheets
                copy = globals()['geo_wb_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).options(numbers=int).value
                # Paste to Geo worksheets
                globals()['ww_geo_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).value = copy
                # Center Allignment
                globals()['ww_geo_{0}'.format(i) + '_ws_{0}'.format(j)].range(range_load).api.HorizontalAlignment = -4108    
        
        # manual copy (roadmap range_setup_b)        
        for i in range(len(select_market)) : 
            
            copy_manual = globals()['geo_wb_{0}'.format(i) + '_ws_2'].range(roadmap_setup_b).options(numbers=int).value
            globals()['ww_geo_{0}'.format(i) + '_ws_2'].range(roadmap_setup_b).value = copy_manual
            globals()['ww_geo_{0}'.format(i) + '_ws_2'].range(range_load).api.HorizontalAlignment = -4108    
            print("\nCopying from " + 'geo_wb_{0}'.format(i) + '_ws_2' + " to " + 'ww_mar_{0}'.format(i) + '_ws_2')
            logging.info("Copying from " + 'geo_wb_{0}'.format(i) + '_ws_2' + " to " + 'ww_mar_{0}'.format(i) + '_ws_2')
        
        print(colored.fg('green'))
        print("\nCompleted copying worksheets from Market to Geo")
        print("\nCompleted running Main Script")
        print(colored.fg('white'))
        logging.info("Completed copying worksheets from Geo to WW")
        logging.info("Completed running Main Script")
        
        #### MAIN 005 : Save and Close
         
        print("\nSaving all worksheets and close the workbook")
        logging.info("Saving all worksheets and close the workbook")
        
        # Set Directory Path for the first market template and save
        globals()['geo_wb_0'].save(call_wb[0])
        print("\nSaving Workbook " + 'geo_wb_0' )
        globals()['geo_wb_0'].close()
        print("\nClosing Workbook " + 'geo_wb_0')
        
        # Save the rest market templates to the path
        for i in range(1, len(select_market)) : 
            print("\nSaving Workbook " + 'geo_wb_{0}'.format(i))
            logging.info("Saving Workbook " + 'geo_wb_{0}'.format(i))
            globals()['geo_wb_{0}'.format(i)].save()
            print("\nClosing Workbook " + 'geo_wb_{0}'.format(i) )
            logging.info("Closing Workbook " + 'geo_wb_{0}'.format(i))
            globals()['geo_wb_{0}'.format(i)].close()
        
        # Save Geo Template
        geo_wb.save(path_geo)
        geo_wb.save()
        geo_wb.close() 
        
        app.quit()
         
        print(colored.fg('green'))
        print("Completed running all scripts. Please check the workbook")
        print(colored.fg('white'))
        logging.info("Completed running all scripts. Please check the workbook")
        
        
    except Exception as e:
        print("Exception : " + e.__str__())
        logging.error("Exception : " + e.__str__())
        app.quit()   




def run_model():
        
    logging.info("Running CATC")              
    
    set_Log_Directory()    

    key_input, login_result  = login()
    
    if ( key_input.__str__() != "None" and login_result == True):    
        
        if (key_input.upper() in get_GEO()):             
            
            if  validate_geo_file_path([key_input]) and validate_market_file_path(get_market(key_input)) :
                run_market_to_GEO(key_input.upper())        
        
        elif(key_input.upper() in get_WW()):
            if  validate_geo_file_path(get_GEO()) and validate_ww_file_path([get_WW(key_input)]) :
                run_GEO_to_Worldwide(key_input.upper())
            
        else:
            
            print("Your Selection is Invalid. Please try again...!")
      
    #======== Last Line of the Code ================================
    input("Press Enter to exit...")

if __name__ == '__main__':
    run_model()