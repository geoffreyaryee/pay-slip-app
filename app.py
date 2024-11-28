def paye(bassic):
    tax=0
    if bassic<=490:
        return tax
    if bassic > 490:
        taxable_income= min(bassic-490,110)
        tax+=taxable_income*0.05
    if bassic > 600:
        taxable_income= min(bassic-600,130)
        tax+=taxable_income*0.10
    if bassic > 730:
        taxable_income= min(bassic-730,3166.67)
        tax+=taxable_income*0.175
    
    return tax

pay=paye(2897.88)

print(pay)