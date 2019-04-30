from bs4 import BeautifulSoup
from openpyxl import Workbook



def suite_parse():
    suite_tag_list = soup.find_all('testsuite')
    suites = {}
    if len(suite_tag_list) == 1:
        name = suite_tag_list[0]['name']
        suites[name]=case_parse(suite_tag_list[0])
    else:
        for suite_tag in suite_tag_list[1:]:
            name = suite_tag['name']
            suites[name]=case_parse(suite_tag)
    return suites

def case_parse(suite_tag):
    cases = {}
    for case_tag in suite_tag.find_all('testcase'):
        steps = steps_parse(case_tag)
        results = results_parse(case_tag)
        try:
            summary_tmp = case_tag.find('summary').text
            summary = BeautifulSoup(summary_tmp,'html.parser').get_text()
        except:
            summary = case_tag.find('summary').text
        #先确定casename不可能为空
        cases[case_tag['name']]={'summary':summary,'steps':steps,'results':results}
    return cases

def steps_parse(case_tag):
    steps = []
    for step in case_tag.find_all('actions'):
        try:
            tmp = step.text
            step_text = BeautifulSoup(tmp,'html.parser').get_text()
        except:
            steps.append(step.text)
        else:
            steps.append(step_text)
        
    return steps

def results_parse(case_tag):
    results = []
    for result in case_tag.find_all('expectedresults'):
        try:
            tmp = result.text
            result_text = BeautifulSoup(tmp,'html.parser').get_text()
        except:
            results.append(result.text)
        else:
            results.append(result_text)

    return results

def w_excel():
    xml_dict = suite_parse()
    i = 1
    for suite_k,suite_v in xml_dict.items():
        sheet['a'+str(i)] = suite_k
        for case_k,case_v in suite_v.items():
            sheet['b'+str(i)] = case_k
            sheet['c'+str(i)] = case_v['summary']
            z = len(case_v['steps'])
            for x in range(z):
                sheet['d'+str(i+x)]= case_v['steps'][x]
                sheet['e'+str(i+x)]= case_v['results'][x]
            i = i + z

if __name__ == '__main__':
    with open('test.xml','r',encoding='utf-8') as f:
        xml = f.read()

    soup = BeautifulSoup(xml,'html.parser')
    
    wb = Workbook()
    sheet = wb.active
    w_excel()
    wb.save('xml.xlsx')
    print(suite_parse())