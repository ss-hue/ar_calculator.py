import PySimpleGUI as sg
import pandas as pd

#Author: Sergio Roldán Salmerón
#Description: Simple Abandon Rate Analyzer that works with the 519: "skills_summary" file from InContact.
#Important Note: The "**" symbols that wrap the rates in the column "Overall Abnd Rate", acts as an alert or indicator meaning the Focus maximum rate was overpassed.


sg.theme('DarkAmber')

layout = [ [sg.Text('Select a file: '), sg.FileBrowse()],
          [sg.Button('Ok'), sg.Button('Cancel')]]

window = sg.Window('AR Analyzer 519', layout)
  


def file_handler(file_path):
    file_ext = file_path.split('.')[1]
    
    if file_ext == "xlsx":
        file = pd.read_excel(file_path)
        abandon_per_skill_dict = dict(zip(file['skill_name'], list(zip(file['Abandon_Cnt'], file['In_SLA'], file['Out_SLA']))))
        return abandon_per_skill_dict

    elif file_ext == "csv":
        file = pd.read_csv(file_path)
        abandon_per_skill_dict = dict(zip(file['skill_name'], list(zip(file['Abandon_Cnt'], file['In_SLA'], file['Out_SLA']))))
        return abandon_per_skill_dict
    else:
        sg.Popup("Please, select a .csv or .xlsx type of file!") 

def higlight_background(ab_rate, focus):
    rate = float(ab_rate.split(" ")[0])
    if(rate > focus):
        return "**{}**".format(ab_rate)
    else:
        return ab_rate
    

def display_calculator(data_per_market):
    sg.theme('LightGrey6')
    sub_headings = ['Type', 'Market','B2B Out SLA', 'B2B In SLA', 'B2B Abnd Calls', 'B2B SLA %', 'B2C Out SLA','B2C In SLA', 'B2C Abnd Calls', 'B2C SLA %', 'Overall Out SLA', 'Overall In SLA', 'Overall Abnd Calls', 'Overall Abnd Rate', 'Focus', 'Overall SLA %']
    layout_calculator = [
        [sg.Table(values=data_per_market,
                  headings=sub_headings, 
                  auto_size_columns=False,
                  max_col_width=13,
                  def_col_width=13,
                  justification='center',
                  key='_TABLE_',
                  row_height=20)],
        [sg.Text('                                                                                                                                                                                                                  By Sergio Roldán Salmerón')]
        ]
    
    window_calculator = sg.Window('AR Analyzer 519', layout_calculator, auto_size_text=True, auto_size_buttons=True, grab_anywhere=False, resizable=True, finalize=True)
    window_calculator['_TABLE_'].expand(True, True)
    window_calculator['_TABLE_'].table_frame.pack(expand=True, fill='both')
    while True:
        event, values = window_calculator.read()
        if event == sg.WIN_CLOSED or event == "Cancel":
            break
        
    window_calculator.close()

def calculator(file_path):
    dach_b2c_abandon_cnt, dach_b2c_inSLA, dach_b2c_outSLA  = ([] for i in range(3))
    dach_b2b_abandon_cnt, dach_b2b_inSLA, dach_b2b_outSLA  = ([] for i in range(3))
    benelux_b2c_abandon_cnt, benelux_b2c_inSLA, benelux_b2c_outSLA  = ([] for i in range(3))
    benelux_b2b_abandon_cnt, benelux_b2b_inSLA, benelux_b2b_outSLA  = ([] for i in range(3))
    france_b2c_abandon_cnt, france_b2c_inSLA, france_b2c_outSLA  = ([] for i in range(3))
    france_b2b_abandon_cnt, france_b2b_inSLA, france_b2b_outSLA  = ([] for i in range(3))
    uki_b2c_abandon_cnt, uki_b2c_inSLA, uki_b2c_outSLA  = ([] for i in range(3))
    uki_b2b_abandon_cnt, uki_b2b_inSLA, uki_b2b_outSLA  = ([] for i in range(3))
    iig_b2c_abandon_cnt, iig_b2c_inSLA, iig_b2c_outSLA  = ([] for i in range(3))
    iig_b2b_abandon_cnt, iig_b2b_inSLA, iig_b2b_outSLA  = ([] for i in range(3))
    iberia_b2c_abandon_cnt, iberia_b2c_inSLA, iberia_b2c_outSLA  = ([] for i in range(3))
    iberia_b2b_abandon_cnt, iberia_b2b_inSLA, iberia_b2b_outSLA  = ([] for i in range(3))
    ee_b2c_abandon_cnt, ee_b2c_inSLA, ee_b2c_outSLA  = ([] for i in range(3))
    ee_b2b_abandon_cnt, ee_b2b_inSLA, ee_b2b_outSLA  = ([] for i in range(3))
    metap_b2c_abandon_cnt, metap_b2c_inSLA, metap_b2c_outSLA  = ([] for i in range(3))
    metap_b2b_abandon_cnt, metap_b2b_inSLA, metap_b2b_outSLA  = ([] for i in range(3))
    sla_rate_dach_b2c = str(0) + " %"
    sla_rate_dach_b2b = str(0) + " %"
    sla_rate_benelux_b2c = str(0) + " %"
    sla_rate_benelux_b2b = str(0) + " %"
    sla_rate_france_b2c = str(0) + " %"
    sla_rate_france_b2b = str(0) + " %"    
    sla_rate_uki_b2c = str(0) + " %"
    sla_rate_uki_b2b = str(0) + " %"  
    sla_rate_iig_b2c = str(0) + " %"
    sla_rate_iig_b2b = str(0) + " %"   
    sla_rate_iberia_b2c = str(0) + " %"
    sla_rate_iberia_b2b = str(0) + " %"   
    sla_rate_ee_b2c = str(0) + " %"
    sla_rate_ee_b2b = str(0) + " %"   
    sla_rate_metap_b2c = str(0) + " %"
    sla_rate_metap_b2b = str(0) + " %"    
                
 
 
 
    
    sk_db_file = pd.read_csv("skill_summary_database.csv")
    abandon_per_skill_dict = file_handler(file_path)
    
    for k,v in abandon_per_skill_dict.items():
        if((sk_db_file['DACH B2C'].dropna()).str.contains(k).any()):
            dach_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            dach_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            dach_b2c_outSLA.append(abandon_per_skill_dict[k][2])
            try:
                sla_rate_dach_b2c = str(round((sum(dach_b2c_inSLA)/(sum(dach_b2c_outSLA) + sum(dach_b2c_inSLA)))*100,2)) + " %"
            except:
                sla_rate_dach_b2c = str(0) + " %"
        if((sk_db_file['DACH B2B'].dropna()).str.contains(k).any()):
            dach_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            dach_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            dach_b2b_outSLA.append(abandon_per_skill_dict[k][2])
            try:
                sla_rate_dach_b2b = str(round((sum(dach_b2b_inSLA)/(sum(dach_b2b_outSLA) + sum(dach_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_dach_b2b = str(0) + " %"
        if((sk_db_file['BENELUX B2C'].dropna()).str.contains(k).any()):
            benelux_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            benelux_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            benelux_b2c_outSLA.append(abandon_per_skill_dict[k][2]) 
            try:
                sla_rate_benelux_b2c = str(round((sum(benelux_b2c_inSLA)/(sum(benelux_b2c_outSLA) + sum(benelux_b2c_inSLA)))*100,2)) + " %"     
            except:
                sla_rate_benelux_b2c = str(0) + " %"   
        if((sk_db_file['BENELUX B2B'].dropna()).str.contains(k).any()):
            benelux_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            benelux_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            benelux_b2b_outSLA.append(abandon_per_skill_dict[k][2]) 
            try:
                sla_rate_benelux_b2b = str(round((sum(benelux_b2b_inSLA)/(sum(benelux_b2b_outSLA) + sum(benelux_b2b_inSLA)))*100,2)) + " %"  
            except:
                sla_rate_benelux_b2b = str(0) + " %"
                          
        if((sk_db_file['FRANCE B2C'].dropna()).str.contains(k).any()):
            france_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            france_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            france_b2c_outSLA.append(abandon_per_skill_dict[k][2]) 
            try:
                sla_rate_france_b2c = str(round((sum(france_b2c_inSLA)/(sum(france_b2c_outSLA) + sum(france_b2c_inSLA)))*100,2)) + " %"       
            except:
                sla_rate_france_b2c = str(0) + " %"
        if((sk_db_file['FRANCE B2B'].dropna()).str.contains(k).any()):
            france_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            france_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            france_b2b_outSLA.append(abandon_per_skill_dict[k][2]) 
            try:
                sla_rate_france_b2b = str(round((sum(france_b2b_inSLA)/(sum(france_b2b_outSLA) + sum(france_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_france_b2b = str(0) + " %"
        if((sk_db_file['UK&I B2C'].dropna()).str.contains(k).any()):
            uki_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            uki_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            uki_b2c_outSLA.append(abandon_per_skill_dict[k][2])  
            try:
                sla_rate_uki_b2c = str(round((sum(uki_b2c_inSLA)/(sum(uki_b2c_outSLA) + sum(uki_b2c_inSLA)))*100,2)) + " %"
            except:
                sla_rate_uki_b2c = str(0) + " %"
        if((sk_db_file['UK&I B2B'].dropna()).str.contains(k).any()):
            uki_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            uki_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            uki_b2b_outSLA.append(abandon_per_skill_dict[k][2])
            try:  
                sla_rate_uki_b2b = str(round((sum(uki_b2b_inSLA)/(sum(uki_b2b_outSLA) + sum(uki_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_uki_b2b = str(0) + " %"
        if((sk_db_file['IIG B2C'].dropna()).str.contains(k).any()):
            iig_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            iig_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            iig_b2c_outSLA.append(abandon_per_skill_dict[k][2])  
            try:
                sla_rate_iig_b2c = str(round((sum(iig_b2c_inSLA)/(sum(iig_b2c_outSLA) + sum(iig_b2c_inSLA)))*100,2)) + " %"
            except:
                sla_rate_iig_b2c = str(0) + " %"
        if((sk_db_file['IIG B2B'].dropna()).str.contains(k).any()):
            iig_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            iig_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            iig_b2b_outSLA.append(abandon_per_skill_dict[k][2]) 
            try: 
                sla_rate_iig_b2b = str(round((sum(iig_b2b_inSLA)/(sum(iig_b2b_outSLA) + sum(iig_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_iig_b2b = str(0) + " %"
        if((sk_db_file['IBERIA B2C'].dropna()).str.contains(k).any()):
            iberia_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            iberia_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            iberia_b2c_outSLA.append(abandon_per_skill_dict[k][2])
            try: 
                sla_rate_iberia_b2c = str(round((sum(iberia_b2c_inSLA)/(sum(iberia_b2c_outSLA) + sum(iberia_b2c_inSLA)))*100,2)) + " %" 
            except:
                sla_rate_iberia_b2c = str(0) + " %"
        if((sk_db_file['IBERIA B2B'].dropna()).str.contains(k).any()):
            iberia_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            iberia_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            iberia_b2b_outSLA.append(abandon_per_skill_dict[k][2])
            try:
                sla_rate_iberia_b2b = str(round((sum(iberia_b2b_inSLA)/(sum(iberia_b2b_outSLA) + sum(iberia_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_iberia_b2b = str(0) + " %"
        if((sk_db_file['EE B2C'].dropna()).str.contains(k).any()):
            ee_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            ee_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            ee_b2c_outSLA.append(abandon_per_skill_dict[k][2]) 
            try:
                sla_rate_ee_b2c = str(round((sum(ee_b2c_inSLA)/(sum(ee_b2c_outSLA) + sum(ee_b2c_inSLA)))*100,2)) + " %" 
            except:
                sla_rate_ee_b2c = str(0) + " %"
        if((sk_db_file['EE B2B'].dropna()).str.contains(k).any()):
            ee_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            ee_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            ee_b2b_outSLA.append(abandon_per_skill_dict[k][2])
            try:
                sla_rate_ee_b2b = str(round((sum(ee_b2b_inSLA)/(sum(ee_b2b_outSLA) + sum(ee_b2b_inSLA)))*100,2)) + " %" 
            except:
                sla_rate_ee_b2b = str(0) + " %"
        if((sk_db_file['METAP B2C'].dropna()).str.contains(k).any()):
            metap_b2c_abandon_cnt.append(abandon_per_skill_dict[k][0])
            metap_b2c_inSLA.append(abandon_per_skill_dict[k][1])
            metap_b2c_outSLA.append(abandon_per_skill_dict[k][2])
            try:  
                sla_rate_metap_b2c = str(round((sum(metap_b2c_inSLA)/(sum(metap_b2c_outSLA) + sum(metap_b2c_inSLA)))*100,2)) + " %"
            except:
                sla_rate_metap_b2c = str(0) + " %"
        if((sk_db_file['METAP B2B'].dropna()).str.contains(k).any()):
            metap_b2b_abandon_cnt.append(abandon_per_skill_dict[k][0])
            metap_b2b_inSLA.append(abandon_per_skill_dict[k][1])
            metap_b2b_outSLA.append(abandon_per_skill_dict[k][2])
            try: 
                sla_rate_metap_b2b = str(round((sum(metap_b2b_inSLA)/(sum(metap_b2b_outSLA) + sum(metap_b2b_inSLA)))*100,2)) + " %"
            except:
                sla_rate_metap_b2b = str(0) + " %"


    dach_overall_abandon_cnt = sum(dach_b2b_abandon_cnt) + sum(dach_b2c_abandon_cnt)
    dach_overall_outSLA = sum(dach_b2b_outSLA) + sum(dach_b2c_outSLA)
    dach_overall_inSLA = sum(dach_b2b_inSLA) + sum(dach_b2c_inSLA)
    try:
        sla_rate_dach_overall = str(round((dach_overall_inSLA/(dach_overall_outSLA + dach_overall_inSLA))*100,2)) + " %"
        dach_overall_abd_rate = str(round((dach_overall_abandon_cnt/(dach_overall_inSLA + dach_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_dach_overall = str(0) + " %"
        dach_overall_abd_rate = str(0) + " %"
    
    benelux_overall_abandon_cnt = sum(benelux_b2b_abandon_cnt) + sum(benelux_b2c_abandon_cnt)
    benelux_overall_outSLA = sum(benelux_b2b_outSLA) + sum(benelux_b2c_outSLA)
    benelux_overall_inSLA = sum(benelux_b2b_inSLA) + sum(benelux_b2c_inSLA)
    try:
        sla_rate_benelux_overall = str(round((benelux_overall_inSLA/(benelux_overall_outSLA + benelux_overall_inSLA))*100,2)) + " %"
        benelux_overall_abd_rate = str(round((benelux_overall_abandon_cnt/(benelux_overall_inSLA + benelux_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_benelux_overall = str(0) + " %"
        benelux_overall_abd_rate = str(0) + " %"    
    
    france_overall_abandon_cnt = sum(france_b2b_abandon_cnt) + sum(france_b2c_abandon_cnt)
    france_overall_outSLA = sum(france_b2b_outSLA) + sum(france_b2c_outSLA)
    france_overall_inSLA = sum(france_b2b_inSLA) + sum(france_b2c_inSLA)
    try:
        sla_rate_france_overall = str(round((france_overall_inSLA/(france_overall_outSLA + france_overall_inSLA))*100,2)) + " %"
        france_overall_abd_rate = str(round((france_overall_abandon_cnt/(france_overall_inSLA + france_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_france_overall = str(0) + " %"
        france_overall_abd_rate = str(0) + " %"      
    
    uki_overall_abandon_cnt = sum(uki_b2b_abandon_cnt) + sum(uki_b2c_abandon_cnt)
    uki_overall_outSLA = sum(uki_b2b_outSLA) + sum(uki_b2c_outSLA)
    uki_overall_inSLA = sum(uki_b2b_inSLA) + sum(uki_b2c_inSLA)
    try:
        sla_rate_uki_overall = str(round((uki_overall_inSLA/(uki_overall_outSLA + uki_overall_inSLA))*100,2)) + " %"
        uki_overall_abd_rate = str(round((uki_overall_abandon_cnt/(uki_overall_inSLA + uki_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_uki_overall = str(0) + " %"
        uki_overall_abd_rate = str(0) + " %"          
    
    iig_overall_abandon_cnt = sum(iig_b2b_abandon_cnt) + sum(iig_b2c_abandon_cnt)
    iig_overall_outSLA = sum(iig_b2b_outSLA) + sum(iig_b2c_outSLA)
    iig_overall_inSLA = sum(iig_b2b_inSLA) + sum(iig_b2c_inSLA)
    try:
        sla_rate_iig_overall = str(round((iig_overall_inSLA/(iig_overall_outSLA + iig_overall_inSLA))*100,2)) + " %"
        iig_overall_abd_rate = str(round((iig_overall_abandon_cnt/(iig_overall_inSLA + iig_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_iig_overall = str(0) + " %"
        iig_overall_abd_rate = str(0) + " %"         
    
    iberia_overall_abandon_cnt = sum(iberia_b2b_abandon_cnt) + sum(iberia_b2c_abandon_cnt)
    iberia_overall_outSLA = sum(iberia_b2b_outSLA) + sum(iberia_b2c_outSLA)
    iberia_overall_inSLA = sum(iberia_b2b_inSLA) + sum(iberia_b2c_inSLA)
    try:
        sla_rate_iberia_overall = str(round((iberia_overall_inSLA/(iberia_overall_outSLA + iberia_overall_inSLA))*100,2)) + " %"
        iberia_overall_abd_rate = str(round((iberia_overall_abandon_cnt/(iberia_overall_inSLA + iberia_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_iberia_overall = str(0) + " %"
        iberia_overall_abd_rate = str(0) + " %"         
    
        
    ee_overall_abandon_cnt = sum(ee_b2b_abandon_cnt) + sum(ee_b2c_abandon_cnt)
    ee_overall_outSLA = sum(ee_b2b_outSLA) + sum(ee_b2c_outSLA)
    ee_overall_inSLA = sum(ee_b2b_inSLA) + sum(ee_b2c_inSLA)
    try:
        sla_rate_ee_overall = str(round((ee_overall_inSLA/(ee_overall_outSLA + ee_overall_inSLA))*100,2)) + " %" 
        ee_overall_abd_rate = str(round((ee_overall_abandon_cnt/(ee_overall_inSLA + ee_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_ee_overall = str(0) + " %"
        ee_overall_abd_rate = str(0) + " %"         
    
            
    metap_overall_abandon_cnt = sum(metap_b2b_abandon_cnt) + sum(metap_b2c_abandon_cnt)
    metap_overall_outSLA = sum(metap_b2b_outSLA) + sum(metap_b2c_outSLA)
    metap_overall_inSLA = sum(metap_b2b_inSLA) + sum(metap_b2c_inSLA)     
    try:
        sla_rate_metap_overall = str(round((metap_overall_inSLA/(metap_overall_outSLA + metap_overall_inSLA))*100,2)) + " %" 
        metap_overall_abd_rate = str(round((metap_overall_abandon_cnt/(metap_overall_inSLA + metap_overall_outSLA))*100,2)) + " %"
    except:
        sla_rate_metap_overall = str(0) + " %"
        metap_overall_abd_rate = str(0) + " %"       
    
    
    
    data_per_market = [
        ['', 'DACH', sum(dach_b2b_outSLA), sum(dach_b2b_inSLA), sum(dach_b2b_abandon_cnt), sla_rate_dach_b2b, sum(dach_b2c_outSLA), sum(dach_b2c_inSLA), sum(dach_b2c_abandon_cnt), sla_rate_dach_b2c, dach_overall_outSLA, dach_overall_inSLA, dach_overall_abandon_cnt, higlight_background(dach_overall_abd_rate,10), "10 %", sla_rate_dach_overall],
        ['H', 'BENELUX', sum(benelux_b2b_outSLA), sum(benelux_b2b_inSLA), sum(benelux_b2b_abandon_cnt), sla_rate_benelux_b2b, sum(benelux_b2c_outSLA), sum(benelux_b2c_inSLA), sum(benelux_b2c_abandon_cnt), sla_rate_benelux_b2c, benelux_overall_outSLA, benelux_overall_inSLA, benelux_overall_abandon_cnt, higlight_background(benelux_overall_abd_rate,10), "10 %", sla_rate_benelux_overall],
        ['V', 'FRANCE', sum(france_b2b_outSLA), sum(france_b2b_inSLA), sum(france_b2b_abandon_cnt), sla_rate_france_b2b, sum(france_b2c_outSLA), sum(france_b2c_inSLA), sum(france_b2c_abandon_cnt), sla_rate_france_b2c, france_overall_outSLA, france_overall_inSLA, france_overall_abandon_cnt, higlight_background(france_overall_abd_rate,10), "10 %", sla_rate_france_overall],
        ['', 'UK&I', sum(uki_b2b_outSLA), sum(uki_b2b_inSLA), sum(uki_b2b_abandon_cnt), sla_rate_uki_b2b, sum(uki_b2c_outSLA), sum(uki_b2c_inSLA), sum(uki_b2c_abandon_cnt), sla_rate_uki_b2c, uki_overall_outSLA, uki_overall_inSLA, uki_overall_abandon_cnt, higlight_background(uki_overall_abd_rate,10), "10 %", sla_rate_uki_overall],
        [ '',  '',  '',  '',  '',  '',  '',  '', '', ''],
        ['', 'IIG', sum(iig_b2b_outSLA), sum(iig_b2b_inSLA), sum(iig_b2b_abandon_cnt), sla_rate_iig_b2b, sum(iig_b2c_outSLA), sum(iig_b2c_inSLA), sum(iig_b2c_abandon_cnt), sla_rate_iig_b2c, iig_overall_outSLA, iig_overall_inSLA, iig_overall_abandon_cnt, higlight_background(iig_overall_abd_rate,10), "10 %", sla_rate_iig_overall],
        ['L', 'IBERIA', sum(iberia_b2b_outSLA), sum(iberia_b2b_inSLA), sum(iberia_b2b_abandon_cnt), sla_rate_iberia_b2b, sum(iberia_b2c_outSLA), sum(iberia_b2c_inSLA), sum(iberia_b2c_abandon_cnt), sla_rate_iberia_b2c, iberia_overall_outSLA, iberia_overall_inSLA, iberia_overall_abandon_cnt, higlight_background(iberia_overall_abd_rate,10), "10 %", sla_rate_iberia_overall],
        ['V', 'EE', sum(ee_b2b_outSLA), sum(ee_b2b_inSLA), sum(ee_b2b_abandon_cnt), sla_rate_ee_b2b, sum(ee_b2c_outSLA), sum(ee_b2c_inSLA), sum(ee_b2c_abandon_cnt), sla_rate_ee_b2c, ee_overall_outSLA, ee_overall_inSLA, ee_overall_abandon_cnt, higlight_background(ee_overall_abd_rate,10), "10 %", sla_rate_ee_overall],
        ['', 'METAP', sum(metap_b2b_outSLA), sum(metap_b2b_inSLA), sum(metap_b2b_abandon_cnt), sla_rate_metap_b2b, sum(metap_b2c_outSLA), sum(metap_b2c_inSLA), sum(metap_b2c_abandon_cnt), sla_rate_metap_b2c, metap_overall_outSLA, metap_overall_inSLA, metap_overall_abandon_cnt, higlight_background(metap_overall_abd_rate,10), "10 %", sla_rate_metap_overall],
    ]
    
    display_calculator(data_per_market)
    

    
    

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel':
        break
    
    file_path = values['Browse']
    
    calculator(file_path)
    

window.close()
