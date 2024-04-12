#Bond Payout Calculator
#Hannah Nevel
#Sep 2023

import os, errno
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
import matplotlib as mpl
import calculations as c
import RegD_Payment_Schedule as sched
import graphs
from matplotlib.font_manager import FontProperties
import datetime as dt
from datetime import datetime
from datetime import date
from matplotlib.ticker import ScalarFormatter
import common_functions as cf

from matplotlib.ticker import NullFormatter  # useful for `logit` scale
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


def Collapsible(layout, key, title='', arrows=(sg.SYMBOL_DOWN, sg.SYMBOL_UP), collapsed=False):
    """
    User Defined Element
    A "collapsable section" element. Like a container element that can be collapsed and brought back
    :param layout:Tuple[List[sg.Element]]: The layout for the section
    :param key:Any: Key used to make this section visible / invisible
    :param title:str: Title to show next to arrow
    :param arrows:Tuple[str, str]: The strings to use to show the section is (Open, Closed).
    :param collapsed:bool: If True, then the section begins in a collapsed state
    :return:sg.Column: Column including the arrows, title and the layout that is pinned
    """
    return sg.Column([[sg.T((arrows[1] if collapsed else arrows[0]), enable_events=True, k=key+'-BUTTON-'),
                       sg.T(title, enable_events=True, key=key+'-TITLE-')],
                      [sg.pin(sg.Column(layout, key=key, visible=not collapsed, metadata=arrows))]], pad=(0,0))

collapse_key = '-dec-'
SYMBOL_DOWN =  '▼'
SYMBOL_UP =    '▲'

#define layout colors
sg.SetOptions (background_color = 'lightgrey',
            element_background_color = 'lightgrey',
            text_element_background_color = 'lightgrey',
               font = ('Arial', 10, 'bold'),
               text_color = 'midnight blue',
               input_text_color ='Black',
               button_color = ('dark goldenrod', 'midnight blue')
               )
#create logo variable
logo = sg.Image("Logo Final - Clear resized.png")

#define inputs for each tab/calculator
invest_title = [[sg.Text('Run Calculator', font=('Arial', 14, 'bold'), text_color='black')]]

investment_input = [[sg.Text('  ')],
    [sg.Text('Initial Investment', size =(15,1), font=('Arial', 11, 'bold')), sg.InputText(size = (13,1),key = '_initial_',default_text='$')],
    [sg.Button('Calculate', key = '_calc_')],
    [sg.Button('Clear Inputs', key = '_clear_')]]

interest_list = ['Compounding','Simple']
term_list = ["1 Year", "3 Year", "5 Year", "7 Year", "9 Year", "11 Year"]

discount_inputs_termsheet =    [[sg.Checkbox('Discounted Bonds?', default=False, key='_discountTF_')],
    [sg.Text('Rate + Lift', size =(12,1)), sg.InputText(size = (9,1),key = '_ratelift_',default_text='')]]

payment_schedule_input1 = [[sg.Text('  ')],
    [sg.Text('Settled Date'),sg.Input( key='_settled_', size=(12,1)), sg.CalendarButton('Choose Date',  target='_settled_', format='%m/%d/%Y', default_date_m_d_y=(1,1,2024), )],
    [sg.Text('Term',size =(12,1)), sg.DropDown(term_list, size =(10,1),key = '_schedterm_')],
    [sg.Text('Interest Type',size =(12,1)),sg.DropDown(interest_list, size =(20,1),key='_interesttype_')],
    [sg.Text("Output Folder:"), sg.Input(key='_outfolder_'), sg.FolderBrowse()],
                  [Collapsible(discount_inputs_termsheet, collapse_key,'Apply Discount', collapsed=False)],
    [sg.Button('Generate Term Sheet', key = '_gensched_')]]
# payment_schedule_input2 =       [ [sg.Text('  ')],
#                                   [sg.Text("Output Folder:"), sg.Input(key='_outfolder_'), sg.FolderBrowse()],
#         [sg.Text('Investor Name', size =(12,1)), sg.InputText(size = (25,1),key = '_name_',default_text='')],
#         [sg.Text('Investor Email', size =(12,1)), sg.InputText(size = (30,1),key = '_email_',default_text='')]]

export_title = [[sg.Text('                       Create a Revenue Statement',font=('Arial', 14, 'bold'), text_color='black')]]

discount_simple_inputs = [[sg.Text('       Lift Inputs', font=('Arial', 11, 'bold'), text_color='midnight blue')],
                [sg.Text('Offering Rate', size =(13,1)), sg.InputText(size = (5,1),key = '_offerrate_',default_text='9%')],
                   [sg.Text('Term', size =(13,1)), sg.InputText(size = (5,1),key = '_term_',default_text='1')],
                   [sg.Text('Rate + Lift', size =(13,1)), sg.InputText(size = (5,1),key = '_bondyield_',default_text='9.0')]]
                    # [sg.Text('Total Bonds', size =(13,1)), sg.InputText(size = (5,1),key = '_sbondtotal_',default_text='100')],
                    # [sg.Checkbox('Check to use Total Bonds to calculate, default is Rate + Lift', default=False, key='_totalbondsTF S_')]]

discount_compound_inputs = [[sg.Text('       Lift Inputs', font=('Arial', 11, 'bold'), text_color='midnight blue')],
                [sg.Text('Offering Rate', size =(13,1)), sg.InputText(size = (5,1),key = '_cofferrate_',default_text='9%')],
                   [sg.Text('Term', size =(13,1)), sg.InputText(size = (5,1),key = '_cterm_',default_text='1')],
                   [sg.Text('Rate + Lift', size =(13,1)), sg.InputText(size = (5,1),key = '_cbondyield_',default_text='9.0')]]
                    # [sg.Text('Total Bonds', size =(13,1)), sg.InputText(size = (5,1),key = '_cbondtotal_',default_text='100')],
                    # [sg.Checkbox('Check to use Total Bonds to calculate, default is Rate + Lift', default=False, key='_totalbondsTF C_')]]

cd_comparison_inputs = [[sg.Text('PHX Rate %', size =(10,1)), sg.InputText(size = (10,1),key = '_phxrate_',default_text='%')],
                        [sg.Text('CD Rate %', size =(10,1)), sg.InputText(size = (10,1),key = '_cdrate_',default_text='%')],
                        [sg.Text('CD Term', size =(10,1)), sg.InputText(size = (10,1),key = '_cdterm_',default_text='')],
                        [sg.Button('Calculate', key = '_cdcalc_')]]


#define outputs for each tab/calculator
output_labels_simple = [
     [sg.Text('Monthly Interest Revenue', size =(21,1))],
     [sg.Text('Annual Interest Revenue', size =(20,1))],
     [sg.Text('Total Interest Accrued', size =(20,1))],
     [sg.Text('Investment Value at Maturity', size =(24,1))]]

output_labels_compounding = [
     [sg.Text('Total Interest Accrued', size =(20,1))],
     [sg.Text('Investment Value at Maturity', size =(25,1))]]

output_labels_discount_col1_s = [[sg.Text('')],
                               [sg.Text('')], 
     [sg.Text('Total Bonds', size =(15,1))],
     [sg.Text('Bond Price', size =(15,1))],
     [sg.Text('Monthly Income', size =(15,1))]]

output_labels_discount_col2_s = [[sg.Text('Investment Breakdown', font=('Arial', 11, 'bold'), text_color='midnight blue')],
     [sg.Text('')],                       
     [sg.Text('Yearly Income', size =(15,1))],
     [sg.Text('Total Interest', size =(15,1))],
     [sg.Text('Lift at Maturity', size =(15,1))]]


output_labels_discount_col3_s = [[sg.Text('')],[sg.Text('Investment Value at Maturity', size =(23,1), font=('Arial', 10, 'bold'))],]

output_labels_discount_col1_c = [[sg.Text('')],
                               [sg.Text('')], 
     [sg.Text('Total Bonds', size =(15,1))],
     [sg.Text('Bond Price', size =(15,1))]]

output_labels_discount_col2_c = [[sg.Text('Investment Breakdown', font=('Arial', 11, 'bold'), text_color='midnight blue')],
     [sg.Text('')],                       
     [sg.Text('Total Interest', size =(15,1))],
     [sg.Text('Lift at Maturity', size =(15,1))]]


output_labels_discount_col3_c = [[sg.Text('')],[sg.Text('Investment Value at Maturity', size =(23,1), font=('Arial', 10, 'bold'))],]

output_labels_cd = [[sg.Text('Monthly Interest Revenue', size =(21,1))],
                    [sg.Text('Annual Interest Revenue', size =(20,1))],
                    [sg.Text('Total Interest Accrued', size =(20,1))],
                    [sg.Text('Investment Value at Maturity', size =(24,1))]]

# output_labels_graph = [[sg.Text('Total Interest Accrued', size =(20,1))],
#                     [sg.Text('Investment Value at Maturity', size =(24,1))]]

#simple interest outputs 
oneyear_output = [[sg.Text('1 Year: 9%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlyone_')],
                  [sg.InputText(size = (14,1),key = '_yearlyone_')],
                  [sg.InputText(size = (14,1),key = '_totalone_')],
                  [sg.InputText(size = (14,1),key = '_matureone_')]]
                                                                
threeyear_output = [[sg.Text('3 Year: 10%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlythree_')],
                  [sg.InputText(size = (14,1),key = '_yearlythree_')],
                  [sg.InputText(size = (14,1),key = '_totalthree_')],
                  [sg.InputText(size = (14,1),key = '_maturethree_')]]

fiveyear_output = [[sg.Text('5 Year: 11%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlyfive_')],
                  [sg.InputText(size = (14,1),key = '_yearlyfive_')],
                  [sg.InputText(size = (14,1),key = '_totalfive_')],
                  [sg.InputText(size = (14,1),key = '_maturefive_')]]

sevenyear_output = [[sg.Text('7 Year: 12%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlyseven_')],
                  [sg.InputText(size = (14,1),key = '_yearlyseven_')],
                  [sg.InputText(size = (14,1),key = '_totalseven_')],
                  [sg.InputText(size = (14,1),key = '_matureseven_')]]

nineyear_output = [[sg.Text('9 Year: 12.5%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlynine_')],
                  [sg.InputText(size = (14,1),key = '_yearlynine_')],
                  [sg.InputText(size = (14,1),key = '_totalnine_')],
                  [sg.InputText(size = (14,1),key = '_maturenine_')]]

elevenyear_output = [[sg.Text('11 Year: 13%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_monthlyeleven_')],
                  [sg.InputText(size = (14,1),key = '_yearlyeleven_')],
                  [sg.InputText(size = (14,1),key = '_totaleleven_')],
                  [sg.InputText(size = (14,1),key = '_matureeleven_')]]

#compounding
compounding_outputs_col1 = [[sg.Text('1 Year: 9%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest1_')],
                  [sg.InputText(size = (14,1),key = '_valuemature1_')]]

compounding_outputs_col2 = [[sg.Text('3 Year: 10%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest3_')],
                  [sg.InputText(size = (14,1),key = '_valuemature3_')]]

compounding_outputs_col3 = [[sg.Text('5 Year: 11%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest5_')],
                  [sg.InputText(size = (14,1),key = '_valuemature5_')]]

compounding_outputs_col4 = [[sg.Text('7 Year: 12%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest7_')],
                  [sg.InputText(size = (14,1),key = '_valuemature7_')]]

compounding_outputs_col5 = [[sg.Text('9 Year: 12.5%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest9_')],
                  [sg.InputText(size = (14,1),key = '_valuemature9_')]]

compounding_outputs_col6 = [[sg.Text('11 Year: 13%', font=('Arial', 11, 'bold'))],
                  [sg.InputText(size = (14,1),key = '_totalinterest11_')],
                  [sg.InputText(size = (14,1),key = '_valuemature11_')]]

#discount
discount_output_col1_s = [[sg.Text('')],
                        [sg.Text('')], 
                  [sg.InputText(size = (12,1),key = '_bonds_')],
                  [sg.InputText(size = (12,1),key = '_bondprice_')],
                  [sg.InputText(size = (12,1),key = '_monthlyreturn_')]]

discount_output_col2_s = [[sg.Text('')],
                        [sg.Text('')], 
                  [sg.InputText(size = (14,1),key = '_yearly_')],
                  [sg.InputText(size = (14,1),key = '_interesttotal_')],
                  [sg.InputText(size = (14,1),key = '_liftvalue_')]]
                  
discount_output_col3_s = [[sg.Text('')],
                  [sg.InputText(size = (14,1), key = '_investvalue_')]]

discount_output_col1_c = [[sg.Text('')],
                        [sg.Text('')], 
                  [sg.InputText(size = (12,1),key = '_cbonds_')],
                  [sg.InputText(size = (12,1),key = '_cbondprice_')]]

discount_output_col2_c = [[sg.Text('')],
                        [sg.Text('')], 
                  [sg.InputText(size = (14,1),key = '_cinteresttotal_')],
                  [sg.InputText(size = (14,1),key = '_cliftvalue_')]]
                  
discount_output_col3_c = [[sg.Text('')],
                  [sg.InputText(size = (14,1),key = '_cinvestvalue_')]]

#comparison tab
cdrate_output = [[sg.Text('      CD', font=('Arial', 11, 'bold'), text_color='midnight blue')],
                  [sg.InputText(size = (11,1),key = '_monthlycd_')],
                  [sg.InputText(size = (11,1),key = '_yearlycd_')],
                  [sg.InputText(size = (11,1),key = '_totalcd_')],
                  [sg.InputText(size = (11,1),key = '_maturecd_')]]
                                                                
phxrate_output = [[sg.Text('   Phoenix', font=('Arial', 11, 'bold'), text_color='midnight blue')],
                  [sg.InputText(size = (11,1),key = '_monthlyphx_')],
                  [sg.InputText(size = (11,1),key = '_yearlyphx_')],
                  [sg.InputText(size = (11,1),key = '_totalphx_')],
                  [sg.InputText(size = (11,1),key = '_maturephx_')]]


#create tab layouts
simple_int_tab_layout = [[sg.Column(output_labels_simple, vertical_alignment='bottom'), 
    sg.Column(oneyear_output,vertical_alignment='center'),
    sg.Column(threeyear_output,vertical_alignment='center'),
    sg.Column(fiveyear_output,vertical_alignment='center'),
    sg.Column(sevenyear_output,vertical_alignment='center'),
    sg.Column(nineyear_output,vertical_alignment='center'),
    sg.Column(elevenyear_output,vertical_alignment='center')]]

comp_int_tab_layout = [
               [sg.Column(output_labels_compounding, vertical_alignment='bottom'),
                sg.Column(compounding_outputs_col1,vertical_alignment='center'),
                sg.Column(compounding_outputs_col2,vertical_alignment='center'),
                sg.Column(compounding_outputs_col3,vertical_alignment='center'),
                sg.Column(compounding_outputs_col4,vertical_alignment='center'),
                sg.Column(compounding_outputs_col5,vertical_alignment='center'),
                sg.Column(compounding_outputs_col6,vertical_alignment='center')]]

discount_s_tab_layout = [
               [sg.Column(discount_simple_inputs, justification='left', vertical_alignment='top'),sg.VSeparator(),
                sg.Column(output_labels_discount_col1_s, vertical_alignment='center'),
                sg.Column(discount_output_col1_s,vertical_alignment='center'),
                sg.Column(output_labels_discount_col2_s, vertical_alignment='center'),
                sg.Column(discount_output_col2_s,vertical_alignment='center'),
                sg.Column(output_labels_discount_col3_s, vertical_alignment='center'),
                sg.Column(discount_output_col3_s, justification='right', vertical_alignment='center')]]

discount_c_tab_layout = [
               [sg.Column(discount_compound_inputs, justification='left', vertical_alignment='bottom'),sg.VSeparator(),
                sg.Column(output_labels_discount_col1_c, vertical_alignment='center'),
                sg.Column(discount_output_col1_c,vertical_alignment='center'),
                sg.Column(output_labels_discount_col2_c, vertical_alignment='center'),
                sg.Column(discount_output_col2_c,vertical_alignment='center'),
                sg.Column(output_labels_discount_col3_c, vertical_alignment='center'),
                sg.Column(discount_output_col3_c,vertical_alignment='center')]]

cd_compare_tab_layout = [
               [sg.Column(cd_comparison_inputs, vertical_alignment='bottom'),sg.VSeparator(),
                sg.Column(output_labels_cd, vertical_alignment='bottom'),
                sg.Column(cdrate_output,vertical_alignment='center'),
                sg.Column(phxrate_output, vertical_alignment='center')]]


#create tab group layout
tabgroup_layout = [[sg.Tab('Simple Interest',simple_int_tab_layout,key='-simple-'),sg.Tab('Compound Interest', comp_int_tab_layout,key='-compound-'),
                         sg.Tab('Lift - Simple',discount_s_tab_layout, key='-simplediscount-'), sg.Tab('Lift - Compound',discount_c_tab_layout, key='-compdiscount-'), \
                            sg.Tab('CD Comparison - Simple',cd_compare_tab_layout,key='-cd-')]]


opened1 = False
#create layout for overall window
layout = [[sg.Column([[logo]], justification='center', expand_x=True, expand_y=True)],
          [sg.Column(invest_title,justification='left',vertical_alignment='top'),sg.Column(export_title,justification='left',vertical_alignment='top')],
          [sg.Column(investment_input, justification='left', vertical_alignment='top'),
           sg.VSeparator(),sg.Column(payment_schedule_input1,justification='left',vertical_alignment='top')],
                    [sg.TabGroup(tabgroup_layout,
                       enable_events=True,
                       key='-tabgroup-')]]

#create window object
window = sg.FlexForm('Reg D Bond Payout Calculator', resizable=True).Layout(layout) 

#loop for calculations, all calculations done in calculations.py and functions are called here to output into window 
while True:
    event, value = window.read() 
    # print(event)
    if event == sg.WINDOW_CLOSED:
        break
    if event == '_calc_': 
              
        #catch program errors for text or blank entry:
        try:
            #tab 1 get values and make updates
            initial = value["_initial_"].replace('$','').replace(',','')
            initial_float = float(initial)
            compare_options_simple = c.all_payouts_regd(initial_float)
                           
            totalinterest_1year = c.insert_commas_and_dollar(float(compare_options_simple[0][2]))
            totalinterest_3year =c.insert_commas_and_dollar(float(compare_options_simple[1][2]))
            totalinterest_5year = c.insert_commas_and_dollar(float(compare_options_simple[2][2]))
            totalinterest_7year = c.insert_commas_and_dollar(float(compare_options_simple[3][2]))
            totalinterest_9year = c.insert_commas_and_dollar(float(compare_options_simple[4][2]))
            totalinterest_11year = c.insert_commas_and_dollar(float(compare_options_simple[5][2]))
            
                           
            window.find_element('_monthlyone_').Update(compare_options_simple[0][0])
            window.find_element('_yearlyone_').Update(compare_options_simple[0][1])
            window.find_element('_totalone_').Update(totalinterest_1year)
            window.find_element('_matureone_').Update(c.value_at_maturity(initial_float,compare_options_simple[0][2]))

            window.find_element('_monthlythree_').Update(compare_options_simple[1][0])
            window.find_element('_yearlythree_').Update(compare_options_simple[1][1])  
            window.find_element('_totalthree_').Update(totalinterest_3year)
            window.find_element('_maturethree_').Update(c.value_at_maturity(initial_float,compare_options_simple[1][2]))

            window.find_element('_monthlyfive_').Update(compare_options_simple[2][0])
            window.find_element('_yearlyfive_').Update(compare_options_simple[2][1])  
            window.find_element('_totalfive_').Update(totalinterest_5year)
            window.find_element('_maturefive_').Update(c.value_at_maturity(initial_float,compare_options_simple[2][2])) 

            window.find_element('_monthlyseven_').Update(compare_options_simple[3][0])
            window.find_element('_yearlyseven_').Update(compare_options_simple[3][1])  
            window.find_element('_totalseven_').Update(totalinterest_7year)
            window.find_element('_matureseven_').Update(c.value_at_maturity(initial_float,compare_options_simple[3][2])) 

            window.find_element('_monthlynine_').Update(compare_options_simple[4][0])
            window.find_element('_yearlynine_').Update(compare_options_simple[4][1])  
            window.find_element('_totalnine_').Update(totalinterest_9year)
            window.find_element('_maturenine_').Update(c.value_at_maturity(initial_float,compare_options_simple[4][2])) 

            window.find_element('_monthlyeleven_').Update(compare_options_simple[5][0])
            window.find_element('_yearlyeleven_').Update(compare_options_simple[5][1])  
            window.find_element('_totaleleven_').Update(totalinterest_11year)
            window.find_element('_matureeleven_').Update(c.value_at_maturity(initial_float,compare_options_simple[5][2]))

            #tab 2 get values and make updates
            compare_options_compound = c.compound_comparisons(initial_float)

            window.find_element('_totalinterest1_').Update(compare_options_compound[0][0])
            window.find_element('_valuemature1_').Update(compare_options_compound[0][1])

            window.find_element('_totalinterest3_').Update(compare_options_compound[1][0])
            window.find_element('_valuemature3_').Update(compare_options_compound[1][1])

            window.find_element('_totalinterest5_').Update(compare_options_compound[2][0])
            window.find_element('_valuemature5_').Update(compare_options_compound[2][1])

            window.find_element('_totalinterest7_').Update(compare_options_compound[3][0])
            window.find_element('_valuemature7_').Update(compare_options_compound[3][1])

            window.find_element('_totalinterest9_').Update(compare_options_compound[4][0])
            window.find_element('_valuemature9_').Update(compare_options_compound[4][1])

            window.find_element('_totalinterest11_').Update(compare_options_compound[5][0])
            window.find_element('_valuemature11_').Update(compare_options_compound[5][1])

            #tab 3 get values and make updates
            offerrate = float(value["_offerrate_"].replace('%',''))
            if '%' in value["_bondyield_"]:
                sbondyield = float(value["_bondyield_"].replace("%",''))
            else:
                sbondyield = float(value["_bondyield_"])
            # if value['_totalbondsTF S_'] == False:
            bondprice = round(c.bond_price(offerrate,float(value["_term_"]),sbondyield),2)
            adjusted_bond_total_s = c.adjusted_total_bonds(initial_float,bondprice)
            
            adjust_total_investment_simple = (adjusted_bond_total_s*1000)
            adjusted_bondprice_simple = round((initial_float/adjust_total_investment_simple)*1000,2)
            totalinterest_phx = c.discount_total_interest(initial_float,float(value["_term_"]),adjusted_bond_total_s,offerrate)
            
            monthlyincome = c.discount_monthly_income(totalinterest_phx,float(value["_term_"]))
            yearlyincome = monthlyincome*12
            lift = c.lift_at_maturity(adjusted_bondprice_simple,adjusted_bond_total_s)
            value_at_mature = c.invest_value_plus_lift(initial_float,totalinterest_phx, lift)
            # else:
            #     adjusted_bond_total = int(value['_sbondtotal_'])
            #     bondprice = round(c.adjusted_total_bonds(initial_float,0,adjusted_bond_total,value['_totalbondsTF S_']),2)
            #     totalinterest_phx = c.discount_total_interest(initial_float,float(value["_term_"]),adjusted_bond_total,offerrate)
                
            #     monthlyincome = c.discount_monthly_income(totalinterest_phx,float(value["_term_"]))
            #     yearlyincome = monthlyincome*12
            #     lift = c.lift_at_maturity(bondprice,adjusted_bond_total)
            #     value_at_mature = c.invest_value_plus_lift(lift,initial_float,totalinterest_phx)                

            window.find_element('_bonds_').Update(adjusted_bond_total_s)
            window.find_element('_bondprice_').Update(c.insert_commas_and_dollar(adjusted_bondprice_simple))
            window.find_element('_monthlyreturn_').Update(c.insert_commas_and_dollar(monthlyincome))

            window.find_element('_yearly_').Update(c.insert_commas_and_dollar(yearlyincome))
            window.find_element('_interesttotal_').Update(c.insert_commas_and_dollar(totalinterest_phx))
            window.find_element('_liftvalue_').Update(c.insert_commas_and_dollar(lift))

            window.find_element('_investvalue_').Update(c.insert_commas_and_dollar(value_at_mature))


            #tab 4 compound lift
            comp_offerrate = float(value["_cofferrate_"].replace('%',''))
            if '%' in value["_cbondyield_"]:
                cbondyield = float(value["_cbondyield_"].replace("%",''))
            else:
                cbondyield = float(value["_cbondyield_"])

            comp_bondprice = c.bond_price(comp_offerrate,float(value["_cterm_"]),cbondyield)
            adjusted_bond_total_c = c.round_up(c.adjusted_total_bonds(initial_float,round(comp_bondprice,2)))
        
            adjust_total_investment_comp = (adjusted_bond_total_c*1000)
            adjusted_bondprice_comp = round((initial_float/adjust_total_investment_comp)*1000,2)
            lift_c = float(adjusted_bond_total_c) * (1000.00 - adjusted_bondprice_comp)
            comp_total_interest = c.compound_interest_total(adjust_total_investment_comp,comp_offerrate,float(value["_cterm_"]))
            interest = float(comp_total_interest[0].replace('$','').replace(',',''))
            comp_value_at_mature = adjust_total_investment_comp + interest
            # else:
            #     adjusted_bond_total_c = int(value['_cbondtotal_'])
            #     comp_bondprice = c.adjusted_total_bonds(initial_float,round(comp_bondprice,2),adjusted_bond_total_c,value['_totalbondsTF C_'])
            #     lift_c = adjusted_bond_total_c * (1000.00 - comp_bondprice)

            #     adjust_total_investment = (adjusted_bond_total_c*1000.00)

            #     comp_total_interest = c.compound_interest_total(adjust_total_investment,comp_offerrate,float(value["_cterm_"]))
            #     interest = float(comp_total_interest[0].replace('$','').replace(',',''))
            #     comp_value_at_mature = adjust_total_investment + lift_c + interest

            window.find_element('_cbonds_').Update(adjusted_bond_total_c)
            window.find_element('_cbondprice_').Update(c.insert_commas_and_dollar(adjusted_bondprice_comp))

            window.find_element('_cinteresttotal_').Update(comp_total_interest[0])
            window.find_element('_cliftvalue_').Update(c.insert_commas_and_dollar(lift_c))

            window.find_element('_cinvestvalue_').Update(c.insert_commas_and_dollar(comp_value_at_mature))

        except ValueError as e:
            
            sg.Popup('Error','Make sure to input required field Initial Investment as a number!') 

    elif event.startswith(collapse_key):
        window[collapse_key].update(visible=not window[collapse_key].visible)
        window[collapse_key+'-BUTTON-'].update(window[collapse_key].metadata[0] if window[collapse_key].visible else window[collapse_key].metadata[1])

        # opened1 = not opened1
        # window['_opendiscount_'].update(SYMBOL_DOWN if opened1 else SYMBOL_UP)
        # window['-dec-'].update(visible=opened1)   

    if event == '_gensched_':
        try:
            #Generating payment schedules based on term
            initial = value["_initial_"].replace('$','').replace(',','')
            initial_float = float(initial)
            mature_value_comp_all = c.compound_comparisons(initial_float)

            inputted_path = value['_outfolder_']
            valid_path = cf.is_pathname_valid(inputted_path)

            if valid_path == True:

                template_path = os.getcwd() + '\\Revenue Statement Template.xlsm'
                #create filename
                todays_date = str(date.today())
                name = 'Revenue Statement - ' + todays_date + '.xlsm'

                output_path = inputted_path + '\\' + name

                unique_outputpath = sched.write_unique_file(output_path)


                schedule_term = int(value["_schedterm_"]. replace(' Year',''))
                
                if schedule_term == 1:
                    schedule_rate = 9
                    compound_mature = mature_value_comp_all[0][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                elif schedule_term == 3: 
                    schedule_rate = 10
                    compound_mature = mature_value_comp_all[1][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                elif schedule_term == 5: 
                    schedule_rate = 11
                    compound_mature = mature_value_comp_all[2][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                elif schedule_term == 7: 
                    schedule_rate = 12
                    compound_mature = mature_value_comp_all[3][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                elif schedule_term == 9: 
                    schedule_rate = 12.5
                    compound_mature = mature_value_comp_all[4][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                elif schedule_term == 11: 
                    schedule_rate = 13
                    compound_mature = mature_value_comp_all[5][1]
                    simple_mature = c.total_interest(initial_float,schedule_rate,schedule_term) + initial_float
                
                rate_lift = value['_ratelift_']
                if rate_lift == '':
                    rate_lift = schedule_rate
                else:
                    rate_lift = float(rate_lift)

                discounted_bonds = value['_discountTF_'] 
                
                interest_type = value['_interesttype_']
                
                settled_date = value['_settled_']

                if discounted_bonds== True:

                    ymax_bprice = round(c.bond_price(13,11,(13+(rate_lift-schedule_rate))),2)

                    ymax_bonds = c.adjusted_total_bonds(initial_float,ymax_bprice)

                    ymax_bprice_adjusted = round((initial_float/(ymax_bonds*1000))*1000,2)

                    ymax_lift = float(ymax_bonds) * (1000.00 - ymax_bprice_adjusted)

                    ymax_interest_str = c.compound_interest_total(ymax_bonds*1000,13,11)

                    ymax_interest = float(ymax_interest_str[0].replace('$','').replace(',',''))

                    ymax_discount= ymax_interest + 5000 + ymax_lift

                    print(ymax_discount)

                    Amount = initial_float * (((1+(13/(100.0*12)))**(12*11)))
                    interest_yaxis_max = Amount - initial_float + 5000

                    if interest_type == 'Simple':
                        bondprice = round(c.bond_price(schedule_rate,schedule_term,rate_lift),2)
                        adjusted_bond_total_s = c.adjusted_total_bonds(initial_float,bondprice)

                        adjust_total_investment_simple = (adjusted_bond_total_s*1000)
                        adjusted_bondprice_simple = round((initial_float/adjust_total_investment_simple)*1000,2)
                        totalinterest_phx = c.discount_total_interest(initial_float,schedule_term,adjusted_bond_total_s,schedule_rate)

                        lift = c.lift_at_maturity(adjusted_bondprice_simple,adjusted_bond_total_s)
                        simple_mature_discounted = c.invest_value_plus_lift(initial_float,totalinterest_phx, lift)

                        simple_payments = sched.simple_interest_schedule_discounted(initial_float,schedule_rate,rate_lift,schedule_term)
                        simple_df = sched.create_payment_table_simple(simple_payments[0], sched.generate_dates(settled_date,schedule_term), simple_payments[1])
                        data_simple = sched.sum_years_s(simple_df)

                        sched.write_to_excel(data_simple, initial_float, simple_mature_discounted, schedule_term, settled_date, interest_type, discounted_bonds, rate_lift, interest_yaxis_max, ymax_discount, unique_outputpath, template_path)
                        
                    elif interest_type == 'Compounding':
                        comp_bondprice = c.bond_price(schedule_rate,schedule_term,rate_lift)
                        adjusted_bond_total_c = c.round_up(c.adjusted_total_bonds(initial_float,round(comp_bondprice,2)))
                        
                        adjust_total_investment_comp = (adjusted_bond_total_c*1000)
                        adjusted_bondprice_comp = round((initial_float/adjust_total_investment_comp)*1000,2)

                        lift_c = float(adjusted_bond_total_c) * (1000.00 - adjusted_bondprice_comp)
                        comp_total_interest = c.compound_interest_total(adjust_total_investment_comp,schedule_rate,schedule_term)
                        interest = float(comp_total_interest[0].replace('$','').replace(',',''))
                        compound_mature_discounted = adjust_total_investment_comp + interest
                        
                        Compound_payments = sched.compound_interest_schedule_discounted(initial_float,schedule_rate,rate_lift,schedule_term)
                        compound_df = sched.create_payment_table_compound(Compound_payments[0], Compound_payments[1],sched.generate_dates(settled_date,schedule_term))
                        data_compound = sched.sum_years_c(compound_df)

                        sched.write_to_excel(data_compound, initial_float, compound_mature_discounted, schedule_term, settled_date, interest_type,  discounted_bonds, rate_lift, interest_yaxis_max, ymax_discount, unique_outputpath, template_path)

                elif discounted_bonds == False:

                    Amount = initial_float * (((1+(13/(100.0*12)))**(12*11)))
                    interest_yaxis_max = Amount - initial_float + 5000
                    if interest_type == 'Simple':
                        total_bonds = c.total_bonds(initial_float)
                        simple_payments = sched.simple_interest_schedule(initial_float,schedule_rate,schedule_term)
                        simple_df = sched.create_payment_table_simple(simple_payments[0], sched.generate_dates(settled_date,schedule_term), simple_payments[1])
                        data_simple = sched.sum_years_s(simple_df)

                        sched.write_to_excel(data_simple, initial_float, c.insert_commas_and_dollar(simple_mature), schedule_term, settled_date, interest_type, discounted_bonds, schedule_rate, interest_yaxis_max, interest_yaxis_max,unique_outputpath, template_path)
                    
                    elif interest_type == 'Compounding':
                        total_bonds = c.total_bonds(initial_float)                    
                        Compound_payments = sched.compound_interest_schedule(initial_float,schedule_rate,schedule_term)
                        compound_df = sched.create_payment_table_compound(Compound_payments[0], Compound_payments[1],sched.generate_dates(settled_date,schedule_term))
                        data_compound = sched.sum_years_c(compound_df)
                    
                        sched.write_to_excel(data_compound, initial_float, compound_mature, schedule_term, settled_date, interest_type,  discounted_bonds, schedule_rate, interest_yaxis_max, interest_yaxis_max,unique_outputpath, template_path)
                
            else:
                sg.Popup('Choose an Output Folder path')

            
        except ValueError as e:
            
            sg.Popup('Error','Make sure to input required fields: Initial Investment, Settled Date, Term, Interest Type!')        

    if event == '_cdcalc_':
        try:

            initial = value["_initial_"].replace('$','').repalce(',','')
            initial_float = float(initial)
            cd_rate = float(value['_cdrate_'].replace('%',''))
            
            cd_term = float(value['_cdterm_'])

            if float(value['_phxrate_'].replace('%','')) in (9, 10, 11, 12, 12.5, 13):
                phx_rate = float(value['_phxrate_'].replace('%',''))
            else:
                sg.Popup('Error','Please enter a valid Phoenix rate offering.')

            compare_options_simple = c.all_payouts_regd(initial_float)
            
            if phx_rate == 9:
                monthly_phx = compare_options_simple[0][0]
                yearly_phx = compare_options_simple[0][1]
                totalinterest_phx = c.insert_commas_and_dollar(float(compare_options_simple[0][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[0][2])

            if phx_rate == 10:
                monthly_phx = compare_options_simple[1][0]
                yearly_phx = compare_options_simple[1][1]                
                totalinterest_phx =c.insert_commas_and_dollar(float(compare_options_simple[1][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[1][2])

            if phx_rate == 11:
                monthly_phx = compare_options_simple[2][0]
                yearly_phx = compare_options_simple[2][1]
                totalinterest_phx = c.insert_commas_and_dollar(float(compare_options_simple[2][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[2][2])

            if phx_rate == 12:
                monthly_phx = compare_options_simple[3][0]
                yearly_phx = compare_options_simple[3][1]
                totalinterest_phx = c.insert_commas_and_dollar(float(compare_options_simple[3][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[3][2])

            if phx_rate == 12.5:
                monthly_phx = compare_options_simple[4][0]
                yearly_phx = compare_options_simple[4][1]
                totalinterest_phx = c.insert_commas_and_dollar(float(compare_options_simple[4][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[4][2])

            if phx_rate == 13:
                monthly_phx = compare_options_simple[5][0]
                yearly_phx = compare_options_simple[5][1]
                totalinterest_phx = c.insert_commas_and_dollar(float(compare_options_simple[5][2]))
                mature_phx = c.value_at_maturity(initial_float,compare_options_simple[5][2])
            
            window.find_element('_monthlyphx_').Update(monthly_phx)
            window.find_element('_yearlyphx_').Update(yearly_phx)
            window.find_element('_totalphx_').Update(totalinterest_phx)
            window.find_element('_maturephx_').Update(mature_phx)

            cd_total = c.cd_total_interest(initial_float, cd_rate, cd_term)

            window.find_element('_monthlycd_').Update(c.cd_monthly_payment(initial_float,cd_rate))
            window.find_element('_yearlycd_').Update(c.cd_annual_payment(initial_float, cd_rate))
            window.find_element('_totalcd_').Update(c.insert_commas_and_dollar(cd_total))
            window.find_element('_maturecd_').Update(c.cd_value_at_maturity(initial_float,cd_total))


        except ValueError as e:
            sg.Popup('Error','Inputs must be a number.') 

    if event == '_clear_':
        window.find_element('_initial_').Update('$')
        window.find_element('_interesttype_').Update('')
        window.find_element('_ratelift_').Update('9.5')
        # window.find_element('_sheetbondtotal_').Update('100')
        window.find_element('_discountTF_').Update(False)
        window.find_element('_settled_').Update('0/0/0000')
        window.find_element('_schedterm_').Update('')
        window.find_element('_outfolder_').Update('')
        window.find_element('_offerrate_').Update('9%')
        window.find_element('_term_').Update('1')
        window.find_element('_bondyield_').Update('9.0')
        window.find_element('_cofferrate_').Update('9%')
        window.find_element('_cterm_').Update('1')
        window.find_element('_cbondyield_').Update('9.0')
        window.find_element('_phxrate_').Update('%')
        window.find_element('_cdterm_').Update('')
        window.find_element('_cdrate_').Update('%')


    elif event == None:
        break


