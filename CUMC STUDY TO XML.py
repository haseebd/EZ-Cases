# EZ Cases for the ASI Imaging System by Haseeb Durrani

# Output XML file name and location
out_xml_file = 'Desktop/C10.30.xml'

# The following variables must be entered by user to generate cases and slide labels specific to the study
study_initials = "A" # What are the study initials?
study_date = "10.30" # What date were the samples put in culture? M.DD Format
protocols = 6 # How many protocols are being tested?
donors = ["Donor1", "Donor2", "Donor3", "Donor4"] # List of Donors - Include date (MDD) if same donor is used for multiple experiments
doses = ["0", "2", "4", "8"] # List of doses
slides_per_dose = 3 # How many slides per dose?

# Optional input parameters
p_id = "123456789" # Patient ID
p_first = "FIRST" # Patient First Name
p_last = "LAST" # Patient Last Name

# After adjusting the following initially, there shouldn't be a need to change them every time
labid = "1" # Lab ID ("1" for CUMC system)
automation_code = "10X RADxScan 1 Circle large hcentred" # Slide Automation Code
tech = "14" # Technology ("14" for ImageDocument scan mode used by CytoRADx currently)



# WARNING: DO NOT CHANGE ANYTHING BELOW UNLESS YOU KNOW WHAT YOU ARE DOING!!!!
# Two lines of code: Case Naming and Slide Naming need to be adjusted based on Case Naming Scheme and Slide Naming Scheme Respectively

from lxml import etree as ET

prots = range(1, protocols + 1)

cases = ET.Element('Cases')

for protocol_number in prots: 
    
    slide_number = 1       
    
    case = ET.SubElement(cases, 'Case')  
    case_name = ET.SubElement(case, 'NAME')
    # Adjust the case naming (below) based on how you want to name the cases. 
    case_name.text = str(study_initials) + str(study_date) + "C" + str(protocol_number)

    #patient_id = ET.SubElement(case, 'PATIENTID')
    #patient_id.text = str(p_id)
    
    #patient_first = ET.SubElement(case, 'PATIENTFIRSTNAME')
    #patient_first.text = str(p_first)
    
    #patient_last = ET.SubElement(case, 'PATIENTLASTNAME')
    #patient_last.text = str(p_last)
    
    lab_id = ET.SubElement(case, 'LAB_ID')
    lab_id.text = str(labid)
    
    slides = ET.SubElement(case, 'Slides')
    
    for donor in donors:
        
        for dose in doses:

            for s in range(1, slides_per_dose + 1):
    
                slide = ET.SubElement(slides, 'Slide')
                
                slide_name = ET.SubElement(slide, 'NAME')
                slide_name.text = str(slide_number)
                
                slide_automation_code = ET.SubElement(slide, 'AUTOMATIONCODE')
                slide_automation_code.text = str(automation_code)
                
                slide_label = ET.SubElement(slide, 'LABEL')
                # The slide label text (below) needs to be adjusted any time there is a change in the slide naming scheme (Example: "C" for Con A # instead of "P" for Protocol #)
                slide_label.text = study_initials + str(donor) + "C" + str(protocol_number) + "G" + str(dose) + str(s) 
                                                                                                                    
                slide_technology = ET.SubElement(slide, 'TECHNOLOGY')
                slide_technology.text = str(tech)

                slide_number += 1 

ET.ElementTree(cases).write(out_xml_file, method = 'xml', pretty_print=True)

with open(out_xml_file, 'r') as f:
    lines = f.readlines()

with open(out_xml_file, 'w') as f:
    f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
    for line in lines:
        f.write(line)
