import json, xlwt

header = ['id', 'subject_id', 'name', 'state_id', 'substate_id', 'community_id', 'city_id',
          'street_id', 'postal', 'fax', 'phone', 'email', 'email_red', 'www', 'officephone',
          'officefax', 'hits', 'url_hits', 'address_before_teryt', 'created', 'modified',
          'teryt_state', 'teryt_substate', 'teryt_community', 'teryt_city', 'teryt_street']


def extract_key_val(line):  # -> (key, val)
    try:
        first_q = 11
        second_q = line[first_q:].index('">')
        return line[first_q:second_q+first_q], line[second_q+13:-7]
    except:
        return '', ''


def lines_to_json(lines):
    json_result = {}
    for e in lines:
        k, v = extract_key_val(e)
        if k in header:
            json_result[k] = v
    return json_result


def create_header():
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('headers')

    for i in range(len(header)):
        sheet.write(0, i, header[i])

    workbook.save('bip-header.xls')


def main():

    xml_path = 'allsubjects.xml'

    lines = open(xml_path, 'r', encoding='utf-8').readlines()

    lines_count = len(lines)
    jsons = []
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('bip.gov.pl')

    print('data to json...')

    a = 0
    # break_up = False

    for i in range(lines_count):
        '''if break_up == True:
            break'''
        if '<row>' in lines[i]:
            for j in range(i+1, lines_count):
                if '</row>' in lines[j]:
                    jsons.append(lines_to_json(lines[i+1:j-1]))
                    i = i + j
                    a += 1
                    '''if a > 10:
                        break_up = True'''
                    break

    print('json to xls')

    for i in range(len(header)):
        sheet.write(0, i, header[i])

    row = 1
    for next_json in jsons:
        row += 1
        for key in next_json:
            sheet.write(row, header.index(key), next_json[key])

    workbook.save('bip.xls')

    print('end')

main()
