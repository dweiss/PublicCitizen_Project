from selenium import webdriver
import xlsxwriter
import os


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
    event_start_beg = driver.find_element_by_name('event_start_beg_dt')
    event_start_end = driver.find_element_by_name('event_start_end_dt')

    # Start Date information
    print('Enter Beginning Start Date Range (##/##/####):')
    # beginning = input()
    print('Enter End Start Date Range (##/##/####):')
    # end = input()
    # 02/14/2021
    event_start_beg.send_keys('02/09/2021')
    event_start_end.send_keys('02/09/2021')

    # Click submit
    search = driver.find_element_by_name('_fuseaction=main.searchresults')
    search.click()

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

    return cases, case_numbers


def filling_sheet(switch, sheet, counter, last_count, cell0, cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8,
                  cell9, cell10, cell16, cell17, cell18, cell19, cell20, cell21):
    i = last_count
    # print('i:', i, 'counter:', counter)
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

def extracting_information(driver, cases, case_numbers):
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
    counter = 0
    last_count = 0
    switch = True
    # Clicking through the files
    for j in cases:
        print('')
        extraction = []
        counter += 1
        # Go to files
        driver.get(j)

        # Yellow Boxes
        questions = driver.find_elements_by_tag_name('th')
        # White Boxes
        answers = driver.find_elements_by_tag_name('td')
        # Seperated by forms
        order = driver.find_elements_by_class_name('aeme')

        #line up list of questions with list of order. Probably will need a 3d list.
        some_counter = 0
        for i in order:
            writing = str(i.__getattribute__('text'))
            for j in questions:
                print(j.__getattribute__('text'))
                print()
                if j.__getattribute__('text') in writing:
                    # print(new_line_remover(j))
                    updated_order = writing.split(j.__getattribute__('text'))
            # for k in answers:
            #     if k.__getattribute__('text') in writing and len(k.__getattribute__('text')) > 0:
            #         updated_order = writing.split(k.__getattribute__('text'))
            print("order:")
            print(updated_order)
            print()


        # Loop Counters
        questions_counter = 0
        answers_counter = 0
        # print(answers.__getattribute__('text'))
        number_of_contaminants = 0

        # for i in answers:
            # print('i:', i.__getattribute__('text'))
        # Gathering Information
        while questions_counter < len(questions):
            print_out = new_line_remover(questions[questions_counter])
            # print('Print Out =', print_out)

            if 'List of Air Contaminant Compounds' in questions[questions_counter].__getattribute__('text'):
                # print('TEST 1', print_out)
                # extraction.append(print_out)
                # extraction.append(' ')
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
                        # print('contaminant', contaminant)
                        # extraction.append(contaminant)
                        sheet.write(counter, 11, contaminant)
                        answers_counter += 1

                        quantity = answers[answers_counter].__getattribute__('text')
                        # extraction.append(quantity)
                        # print('quantity', quantity)
                        answers_counter += 1

                        unit_measurement = answers[answers_counter].__getattribute__('text')
                        # extraction.append(unit_measurement)
                        # print('unit measurement', unit_measurement)
                        sheet.write(counter, 15, unit_measurement)
                        answers_counter += 1

                        counter += 1

                        emissionlimit = answers[answers_counter].__getattribute__('text')
                        # extraction.append(emissionlimit)
                        # print('emissionlimit', emissionlimit)
                        cell16 = float(emissionlimit)
                        # sheet.write(counter, 17, emissionlimit)
                        answers_counter += 1

                        units = answers[answers_counter].__getattribute__('text')
                        # extraction.append(units)
                        # print('units', units)
                        cell17 = units
                        # sheet.write(counter, 18, units)
                        answers_counter += 1

                        authorization = answers[answers_counter].__getattribute__('text')
                        # extraction.append(authorization)
                        cell18 = authorization
                        # print('authorization', authorization)
                        answers_counter += 1

                        number_of_contaminants -= 1
                else:
                    print_out1 = print_out + ' ' + answers[answers_counter].__getattribute__('text')
                    # print('print out =', print_out1)
                    answer = answers[answers_counter].__getattribute__('text')
                    # print('TEST 8', print_out)
                    if print_out == 'Incident Tracking Number:':
                        # sheet.write(counter, 0, int(answer))
                        cell0 = int(answer)
                    elif print_out == 'RN:':
                        # sheet.write(counter, 1, answer)
                        cell1 = answer
                    elif print_out == 'Regulated Entity Name:':
                        # sheet.write(counter, 2, answer)
                        cell2 = answer
                    elif print_out == 'Physical Location:':
                        # sheet.write(counter, 3, answer)
                        cell3 = answer
                    elif print_out == 'County:':
                        # sheet.write(counter, 4, answer)
                        cell4 = answer
                    elif print_out == 'Notification Jurisdictions:':
                        # sheet.write(counter, 5, answer)
                        cell5 = answer
                    elif print_out == 'Date and Time Event Discovered or Scheduled Activity Start:':
                        # sheet.write(counter, 6, answer)
                        cell6 = answer
                    elif print_out == 'Date and Time Event or Scheduled Activity Ended:':
                        # sheet.write(counter, 7, answer)
                        cell7 = answer
                    elif print_out == 'Event/Activity Type:':
                        # sheet.write(counter, 8, answer)
                        cell8 = answer
                    elif print_out == '1 - Emission Point Common Name:':
                        # sheet.write(counter, 9, answer)
                        cell9 = answer
                    elif print_out == 'Emission Point Number:':
                        # sheet.write(counter, 10, answer)
                        cell10 = answer
                    elif print_out == 'Basis Used to Determine Quantities and Any Additional Information Necessary to Evaluate the Event:':
                        # sheet.write(counter, 19, answer)
                        cell19 = answer
                    elif print_out == 'Cause of Emission Event or Excess Opacity Event, or Reason for Scheduled Activity:':
                        # sheet.write(counter, 20, answer)
                        cell20 = answer
                    elif print_out == 'Actions Taken, or Being Taken, to Minimize Emissions And/or Correct the Situation:':
                        # sheet.write(counter, 21, answer)
                        cell21 = answer
                    # print('TEST 9', answer)
                    # extraction.append(answers[answers_counter].__getattribute__('text'))
                    questions_counter += 1
                    answers_counter += 1
        filling_sheet(switch, sheet, counter, last_count, cell0, cell1, cell2, cell3, cell4, cell5, cell6, cell7, cell8,
                      cell9, cell10, cell16, cell17, cell18, cell19, cell20, cell21)
        if switch:
            switch = False
        last_count = counter
    book.close()


def main():
    driver = setup()
    cases, case_numbers = collecting_information(driver)
    extracting_information(driver, cases, case_numbers)


main()
