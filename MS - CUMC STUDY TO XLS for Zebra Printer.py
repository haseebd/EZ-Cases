# EZ Labels for the Zebra Printer by Haseeb Durrani

# Output XLS file name and location
out_xls_file = 'Desktop/C4.29.xls'

# The following variables must be entered by user to generate slide labels specific to the study
study_name = "ConAComp" # What is the study name?
protocols = 6 # How many protocols are being tested?
donors = ["F223X429", "F537429", "M574429", "M599429"] # List of Donors - Include date (MDD) if same donor is used for multiple experiments
doses = ["000", "020", "040", "080"] # List of doses
slides_per_dose = 3 # How many slides per dose?

# WARNING: DO NOT CHANGE ANYTHING BELOW UNLESS YOU KNOW WHAT YOU ARE DOING!!!!
# One line of code: Slide Naming needs to be adjusted based on Slide Naming Scheme

import xlwt

book = xlwt.Workbook(encoding="utf-8")

sheet1 = book.add_sheet("Sheet 1")

style = xlwt.easyxf("font: bold 1; align: horiz center")

sheet1.write(0, 0, "Study Name", style)
sheet1.write(0, 1, "Donor ID", style)
sheet1.write(0, 2, "Condition", style)
sheet1.write(0, 3, "Dose", style)
sheet1.write(0, 4, "Slide Number", style)
sheet1.write(0, 5, "Slide Label", style)
sheet1.write(0, 6, "Slide # in Case", style)

prots = range(1, protocols + 1)

x = 1
c = 1 

for protocol_number in prots: 
    
    slide_number = 1       
    
    for donor in donors:
        
        for dose in doses:

            for s in range(1, slides_per_dose + 1):
                
                sheet1.write(x, 0, str(study_name), xlwt.easyxf("align: horiz right"))
                sheet1.write(x, 1, str(donor), xlwt.easyxf("align: horiz right"))
                sheet1.write(x, 2, "C" + str(protocol_number), xlwt.easyxf("align: horiz right"))
                sheet1.write(x, 3, str(dose), xlwt.easyxf("align: horiz right"))
                sheet1.write(x, 4, s, xlwt.easyxf("align: horiz right"))
                # The slide label text (below) needs to be adjusted any time there is a change in the slide naming scheme (Example: "C" for Con A # instead of "P" for Protocol #)
                sheet1.write(x, 5, study_name + "." + str(donor) + "C" + str(protocol_number) + "D" + str(dose) + "." + str(s), xlwt.easyxf("align: horiz right"))
                sheet1.write(x, 6, slide_number, xlwt.easyxf("align: horiz right"))
                
                x += 1
                slide_number += 1

sheet1.col(0).width = 256 * (15)
sheet1.col(1).width = 256 * (len(max(donors)) + 5)
sheet1.col(2).width = 256 * (10)
sheet1.col(3).width = 256 * (7)
sheet1.col(4).width = 256 * (15)
sheet1.col(5).width = 256 * (30)
sheet1.col(6).width = 256 * (15)


book.save(out_xls_file)