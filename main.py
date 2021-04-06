from selenium import webdriver
import xlsxwriter
import os

global switch
global counter
global last_count


def number_finder(text):
    return any(i.isdigit() for i in text)


def new_line_remover(question):
    if '\n' in question.__getattribute__('text'):
        return question.__getattribute__('text').replace('\n', '')
    else:
        return question.__getattribute__('text')


def setup():
    # Naming Driver
    driver = webdriver.Chrome()

    # Go to Website.
    driver.get('https://www2.tceq.texas.gov/oce/eer/index.cfm')

    # Find input boxes
    print('If you would like a single search enter "Y":')
    single_search = input()
    if single_search == 'Y':
        incident_number = driver.find_element_by_name('incid_track_num')
        print('Enter Incident Number:')
        incident_number_input = input()
        incident_number.send_keys(incident_number_input)
    else:
        event_start_beg = driver.find_element_by_name('event_start_beg_dt')
        event_start_end = driver.find_element_by_name('event_start_end_dt')
        event_end_beg = driver.find_element_by_name('event_end_beg_dt')
        event_end_end = driver.find_element_by_name('event_end_end_dt')
        cn = driver.find_element_by_name('cn_txt')
        customer_name = driver.find_element_by_name('cust_name')
        rn = driver.find_element_by_name('rn_txt')
        regulated_entity_name = driver.find_element_by_name('re_name')
        county = driver.find_element_by_name('ls_cnty_name')
        region = driver.find_element_by_name('ls_region_cd')
        event_type = driver.find_element_by_name('ls_event_typ_cd')

        print('Enter information here, enter nothing to omit.')
        print('Enter Beginning Start Date Range (##/##/####):')
        beginning_start = input()
        event_start_beg.send_keys(beginning_start)
        print('Enter Last Start Date Range (##/##/####):')
        last_start = input()
        event_start_end.send_keys(last_start)
        print('Enter Beginning End Date Range (##/##/#### must be after 1/31/2003):')
        end_start = input()
        event_end_beg.send_keys(end_start)
        print('Enter Last End Date Range (##/##/#### must be after 1/31/2003):')
        end_end = input()
        event_end_end.send_keys(end_end)
        print('Enter CN:')
        cn_input = input()
        cn.send_keys(cn_input)
        print('Enter Customer Name:')
        customer_name_input = input()
        customer_name.send_keys(customer_name_input)
        print('Enter RN:')
        rn_input = input()
        rn.send_keys(rn_input)
        print('Enter Regulated Entity Name:')
        regulated_entity_name_input = input()
        regulated_entity_name.send_keys(regulated_entity_name_input)
        print('Enter County:')
        county_input = input()
        county.send_keys(county_input)
        print('Enter Region (REGION ## - _____):')
        region_input = input()
        region.send_keys(region_input)
        print('Enter Event Type:')
        event_type_input = input()
        event_type.send_keys(event_type_input)


    # Click submit
    search = driver.find_element_by_name('_fuseaction=main.searchresults')
    search.click()
    print('Processing Now...')

    return driver


def collecting_information(driver):
    # Finds the clickable links that are related to the forms
    cases = []
    case_numbers = []
    repeat = True
    while repeat:
        link = driver.find_elements_by_class_name('datadisplay')
        for lines in link:
            case = lines.find_elements_by_tag_name('a')
            for sublinks in case:
                if number_finder(sublinks.__getattribute__('text')):
                    cases.append(sublinks.get_attribute('href'))
                    case_numbers.append(sublinks.__getattribute__('text'))

        navigation = driver.find_element_by_class_name('pagingnav')
        navigation = navigation.find_elements_by_css_selector('a')

        next_impossible = True
        for pages in navigation:
            if pages.get_attribute('text') == '>':
                driver.get(pages.get_attribute('href'))
                next_impossible = False

        if next_impossible:
            repeat = False

    return cases

def time_converter(half, time):
    sep = time.split(":")
    hour = int(sep[0])
    if (hour == 12):
        hour = 0
    if (half == 'PM'):
        hour += 12
    return str(hour) + ':' + str(sep[1])

def filling_sheet(sheet, cell0, cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8,
                  cell9, cell10, cell16, cell17, cell18, cell19, cell20, cell21):
    i = last_count
    if switch:
        i += 1
    while i < counter:
        sheet.write(i, 0, cell0)
        sheet.write(i, 1, cell1)
        sheet.write(i, 2, cell2)
        sheet.write(i, 3, cell3)
        sheet.write(i, 4, cell4)
        sheet.write(i, 5, cell5)
        sheet.write(i, 6, cell6)
        sheet.write(i, 7, cell7)
        sheet.write(i, 8, cell8)
        sheet.write(i, 9, cell9)
        sheet.write(i, 10, cell10)
        sheet.write(i, 16, cell16)
        sheet.write(i, 17, cell17)
        sheet.write(i, 18, cell18)
        sheet.write(i, 19, cell19)
        sheet.write(i, 20, cell20)
        sheet.write(i, 21, cell21)
        i += 1

def contaminants(questions, answers, order, sheet):
    # Loop Counters
    questions_counter = 0
    answers_counter = 0
    number_of_contaminants = 0
    global switch
    global counter
    global last_count

    cell16 = float(0)
    cell17 = ''
    cell18 = ''

    # Gathering Information
    while questions_counter < len(questions):
        print_out = new_line_remover(questions[questions_counter])
        if 'List of Air Contaminant Compounds' in questions[questions_counter].__getattribute__('text'):
            splitting = questions[questions_counter].__getattribute__('text').split()
            for s in splitting:
                if s.isnumeric():
                    number_of_contaminants = int(s)
            questions_counter += 1
        else:
            if number_of_contaminants > 0:
                questions_counter += 6

                while number_of_contaminants > 0:
                    contaminant = (answers[answers_counter].__getattribute__('text'))
                    sheet.write(counter, 11, contaminant)
                    answers_counter += 1

                    quantity = answers[answers_counter].__getattribute__('text')
                    sheet.write(counter, 12, float(quantity))
                    answers_counter += 1

                    unit_measurement = answers[answers_counter].__getattribute__('text')
                    sheet.write(counter, 15, unit_measurement)
                    answers_counter += 1

                    counter += 1

                    emissionlimit = answers[answers_counter].__getattribute__('text')
                    cell16 = float(emissionlimit)
                    # sheet.write(counter, 17, emissionlimit)
                    answers_counter += 1

                    units = answers[answers_counter].__getattribute__('text')
                    cell17 = units
                    # sheet.write(counter, 18, units)
                    answers_counter += 1

                    authorization = answers[answers_counter].__getattribute__('text')
                    cell18 = authorization
                    answers_counter += 1

                    number_of_contaminants -= 1
            else:
                count = True
                answer = answers[answers_counter].__getattribute__('text')
                if print_out == 'Incident Tracking Number:':
                    cell0 = int(answer)
                elif print_out == 'RN:':
                    cell1 = answer
                elif print_out == 'Regulated Entity Name:':
                    cell2 = answer
                elif print_out == 'Physical Location:':
                    cell3 = answer
                elif print_out == 'County:':
                    cell4 = answer
                elif print_out == 'Notification Jurisdictions:':
                    x = answer.split(" ")
                    cell5 = int(x[1])
                elif print_out == 'Process Unit or Area Common Names':
                    text = order[2].__getattribute__('text').split("\n")
                    answers_counter += (len(text) - 1)
                    questions_counter += 1
                    count = False
                elif print_out == 'Facility Common Name':
                    text = order[3].__getattribute__('text').split("\n")
                    answers_counter += (len(text) - 1) * 2
                    questions_counter += 2
                    count = False
                elif print_out == 'Date and Time Event Discovered or Scheduled Activity Start:':
                    y = answer.split(" ")
                    time1 = time_converter(y[-1], y[-2])
                    cell6 = str(y[0]) + ' ' + time1
                elif print_out == 'Date and Time Event or Scheduled Activity Ended:':
                    z = answer.split(" ")
                    time2 = time_converter(z[-1], z[-2])
                    cell7 = str(z[0]) + ' ' + time2
                elif print_out == 'Event/Activity Type:':
                    cell8 = answer
                elif print_out == '1 - Emission Point Common Name:':
                    cell9 = answer
                elif print_out == 'Emission Point Number:':
                    cell10 = answer
                elif print_out == 'Basis Used to Determine Quantities and Any Additional Information Necessary to Evaluate the Event:':
                    cell19 = answer
                elif print_out == 'Cause of Emission Event or Excess Opacity Event, or Reason for Scheduled Activity:':
                    cell20 = answer
                elif print_out == 'Actions Taken, or Being Taken, to Minimize Emissions And/or Correct the Situation:':
                    cell21 = answer
                if count:
                    questions_counter += 1
                    answers_counter += 1
    filling_sheet(sheet, cell0, cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8,
                cell9, cell10, cell16, cell17, cell18, cell19, cell20, cell21)
    if switch:
        switch = False
    last_count = counter


def extracting_information(driver, cases):
    direct = os.getcwd() + '\TCEQ_Data.xlsx'
    book = xlsxwriter.Workbook(direct)
    sheet = book.add_worksheet('Cases')
    # sheet.set_column(0, 1, 100)
    sheet.write(0, 0, 'INCIDENT NO.')
    sheet.write(0, 1, 'RN')
    sheet.write(0, 2, 'RE NAME')
    sheet.write(0, 3, 'PHYSICAL LOCATION')
    sheet.write(0, 4, 'COUNTY')
    sheet.write(0, 5, 'TCEQ REGION')
    sheet.write(0, 6, 'START DATE/TIME')
    sheet.write(0, 7, 'END DATE/TIME')
    sheet.write(0, 8, 'EVENT TYPE')
    sheet.write(0, 9, 'EMISSION POINT NAME')
    sheet.write(0, 10, 'EPN')
    sheet.write(0, 11, 'CONTAMINANT')
    sheet.write(0, 12, 'EST QUANTITY/OPACITY')
    sheet.write(0, 13, 'ESTIMATED IND')
    sheet.write(0, 14, 'AMOUNT UNK IND')
    sheet.write(0, 15, 'UNITS')
    sheet.write(0, 16, 'EMISSION LIMIT')
    sheet.write(0, 17, 'LIMIT UNITS')
    sheet.write(0, 18, 'AUTHORIZATION COMMENT')
    sheet.write(0, 19, 'COMMENT NO.')
    sheet.write(0, 20, 'Cause of Emission Event')
    sheet.write(0, 21, 'Actions Taken')
    global counter
    global last_count
    global switch

    counter = 0
    last_count = 0
    switch = True
    # Clicking through the files
    count = True
    for j in cases:
        if (count):
            counter += 1
            count = False
        # Go to files
        driver.get(j)

        # Yellow Boxes
        questions = driver.find_elements_by_tag_name('th')
        # White Boxes
        answers = driver.find_elements_by_tag_name('td')
        # Seperated by forms
        order = driver.find_elements_by_class_name('aeme')

        contaminants(questions, answers, order, sheet)
    book.close()


def main():
    driver = setup()
    cases = collecting_information(driver)
    extracting_information(driver, cases)


main()
