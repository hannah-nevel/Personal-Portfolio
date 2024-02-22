
import math
import matplotlib
import matplotlib.pyplot as plt
import numpy as np


#Comopund Interest
def compound_interest_total(initial, rate, term):
    Amount = initial * (((1+(rate/(100.0*12)))**(12*term)))
    Total_interest = Amount - initial                    
    return [insert_commas_and_dollar(Total_interest),insert_commas_and_dollar(Amount)]

def compound_comparisons(initial):
    one_year = compound_interest_total(initial,9,1)
    three_year = compound_interest_total(initial,10,3)
    five_year = compound_interest_total(initial,11,5)
    seven_year = compound_interest_total(initial,12,7)
    nine_year = compound_interest_total(initial,12.5,9)
    eleven_year = compound_interest_total(initial,13,11)

    all = [one_year,three_year,five_year,seven_year,nine_year,eleven_year]
    return all

#Simple Interest Reg D
def monthly_payment(initial, rate):
    monthly_payout = round((initial*(rate/100))/12,2)
    return insert_commas_and_dollar(monthly_payout)

def annual_payment(initial, rate):
    annual_payout = round(initial*(rate/100),2)
    return insert_commas_and_dollar(annual_payout)

def total_interest(initial, rate, term):
    total = round((initial*(rate/100))*term,2)
    return total

def value_at_maturity(initial,total):
    maturity_value = float(total) + initial
    return insert_commas_and_dollar(maturity_value)

def all_payouts_regd(initial):
    one_year = [monthly_payment(initial,9),annual_payment(initial,9),total_interest(initial,9,1)]
    three_year = [monthly_payment(initial,10),annual_payment(initial,10),total_interest(initial,10,3)]
    five_year = [monthly_payment(initial,11),annual_payment(initial,11),total_interest(initial,11,5)]
    seven_year = [monthly_payment(initial,12),annual_payment(initial,12),total_interest(initial,12,7)]
    nine_year = [monthly_payment(initial,12.5),annual_payment(initial,12.5),total_interest(initial,12.5,9)]
    eleven_year = [monthly_payment(initial,13),annual_payment(initial,13),total_interest(initial,13,11)]

    all = [one_year,three_year,five_year,seven_year,nine_year,eleven_year]
    
    return all
#Simple Interest Reg A
def all_payouts_rega(initial):
    
    three_year = [monthly_payment(initial,9),annual_payment(initial,9),total_interest(initial,9,3)]
    five_year = [monthly_payment(initial,10),annual_payment(initial,10),total_interest(initial,10,5)]
    

    all = [three_year,five_year]
    
    return all

#Simple Interest Discount

def bond_price(rate, term, bond_yield):
    parvalue = 1000
    coupon_rate = (rate/100)*parvalue
    bond_price = coupon_rate*(1-(1+(bond_yield/100))**(-1*term))/(bond_yield/100)+(parvalue/(1+(bond_yield/100))**term)
    return bond_price

def total_bonds(initial):
    totalbonds = initial/1000
    return totalbonds

def adjusted_total_bonds(initial, bondprice):
    # if use_bond_total:
    #     bondprice = initial/bonds
    #     return bondprice
    # else:
    discounted_bonds = initial/bondprice
    adjusted_total_bonds = math.ceil(discounted_bonds)
    return adjusted_total_bonds

def discount_total_interest(initial,term,adjusted_total_bonds,offerrate):
    parvalue = 1000
    totalinterest = (adjusted_total_bonds*(((offerrate/100)*parvalue)/12))*(term*12)
    return totalinterest

def discount_yearly_income(initial,rate):
    yearly = round(initial*(rate/100),2)
    return insert_commas_and_dollar(yearly)
    
def discount_monthly_income(totalinterst,term):
    monthly = totalinterst/(term*12)
    return monthly

def lift_at_maturity(bondprice,adjustedtotalbonds):
    parvalue = 1000
    lift = adjustedtotalbonds*(parvalue-bondprice)
    return lift

def invest_value_plus_lift(initial,totalinterest):
    totalvalue = initial+totalinterest
    return totalvalue

def return_discount(initial, adjusted_total_bonds, parvalue):
    adjusted_total_investment_facevalue = adjusted_total_bonds*parvalue
    return_discount = round(adjusted_total_investment_facevalue - initial,1)
    return return_discount

#cd simple interest

def cd_monthly_payment(initial, cdrate):
    monthly_payout = round((initial*(cdrate/100))/12,2)
    return insert_commas_and_dollar(monthly_payout)

def cd_annual_payment(initial, cdrate):
    annual_payout = round(initial*(cdrate/100),2)
    return insert_commas_and_dollar(annual_payout)

def cd_total_interest(initial, cdrate, cdterm):
    total = round((initial*(cdrate/100))*cdterm,2)
    return total

def cd_value_at_maturity(initial,total):
    maturity_value = float(total) + initial
    return insert_commas_and_dollar(maturity_value)

#formatting
def insert_commas_and_dollar(num):
    comma_num = '{:,.2f}'.format(num)
    result = '$'+comma_num

    return result

def round_up(n, decimals=0):
    multiplier = 10**decimals
    return math.ceil(n * multiplier) / multiplier

