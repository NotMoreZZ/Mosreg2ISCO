import copy
import re
from time import sleep

import yaml
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as cond
from selenium.webdriver.support.ui import Select, WebDriverWait


class Login(object):
    BYESIAPASS = 0
    BYBOTHPASSES = 1
    BYNOPASSES = 2

    def __init__(self, file, logintype):
        super().__init__()
        self.type = logintype

        with open(file, "rt") as fhandler:
            text = fhandler.read()
            data = yaml.safe_load(text)

        if self.type == Login.BYESIAPASS or self.type == Login.BYBOTHPASSES:
            self.esianame = data['esianame']
            self.esiapass = data['esiapass']
        if self.type == Login.BYBOTHPASSES:
            self.isconame = data['isconame']
            self.iscopass = data['iscopass']


class MosregToISCO(object):
    _SUBJECTS = {
        'алгебра': 'Алгебра',
        'алгебра и начала анализа': 'Алгебра',
        'английский язык': 'Англ. язык',
        'астрономия': 'Астрономия',
        'биология': 'Биология',
        'всеобщая история': 'Всеобщая история',
        'география': 'География',
        'геометрия': 'Геометрия',
        'изобразительное искусство': 'ИЗО',
        'информатика': 'Инф. и ИКТ',
        'искусство (мхк)': 'Искусство (МХК)',
        'история': 'История',
        'история России': 'История России',
        'литература': 'Литература',
        'литературное чтение': 'Лит. чтение',
        'математика': 'Математика',
        'музыка': 'Музыка',
        'немецкий язык': 'Нем. язык',
        'обществознание': 'Обществознание',
        'окружающий мир': 'Окр. мир',
        'основы безопасности жизнедеятельности': 'ОБЖ',
        'основы религиозных культур и светской этики': '',
        'право': 'Право',
        'русский язык': 'Рус. язык',
        'технология': 'Технология',
        'физика': 'Физика',
        'французский язык': 'Франц. язык',
        'физическая культура': 'Физкультура',
        'химия': 'Химия',
        'экономика': 'Экономика'}

    def __init__(self):
        super().__init__()
        self.driver = webdriver.Firefox()

    def itog(self, subject, groupname):
        # Navigate to appropriate class in journal
        classname = re.search('[\d]{1,2}[А-Е]', groupname).group(0)
        self.driver.get('https://schools.school.mosreg.ru/journals/')
        self._find_element_try_hard(
            By.LINK_TEXT, self._mosreg_class_name(classname)).click()
        self.driver.get(
            self._find_element_try_hard(
                By.LINK_TEXT, self._SUBJECTS[subject]).get_property('href'))
        self._find_element_try_hard(By.LINK_TEXT, 'Итоговые').click()
        sleep(5)

        # Get studentsinfo
        # TODO: that's awful and will work only with one quarter/semester
        # TODO: probably need to autodownload excel exports and work with them
        info = []
        students = self._find_elements_try_hard(
            By.XPATH, '//a[@title="Перейти на страницу оценок ученика"]')
        sleep(1)
        ids = self._find_elements_try_hard(
            By.XPATH,
            '//a[@title="Перейти на страницу оценок ученика"]/parent::*')
        sleep(1)
        for num in range(len(students)):
            info.append(
                [students[num].get_attribute('textContent').replace('ё', 'е'),
                 ids[num].get_attribute('id'), 0, 0, 0, 0, 0, 0, 0, 0, 0])

        # Get grades
        # TODO: same as previous TODO block
        gradetextxpatha = '//td[@id="{}"]/descendant::a'
        gradetextxpathspan = '//td[@id="{}"]/descendant::span/descendant::span'
        gradecells = self._find_elements_try_hard(
            By.XPATH, '//div[@class="pres"]/ancestor::td')
        sleep(1)
        for gradecell in gradecells:
            sid = gradecell.get_attribute('id')
            cellclass = gradecell.get_attribute('class')
            for studentinfo in info:
                if not sid.count('_' + studentinfo[1] + '_'):
                    continue
                if cellclass.count('mX'):
                    grade = '-'
                else:
                    if cellclass.count('mark_v'):
                        grade = self._find_element_try_hard(By.XPATH,
                                                            gradetextxpathspan.format(
                                                                sid)).get_attribute('textContent')
                    else:
                        grade = self._find_element_try_hard(By.XPATH,
                                                            gradetextxpatha.format(
                                                                sid)).get_attribute('textContent')
                for field in range(len(studentinfo)):
                    if not studentinfo[field]:
                        studentinfo[field] = grade
                        break

        # Navigate to "Итоговая ведомость" page
        self.driver.get('https://isko.mosreg.ru/control')
        self._find_element_try_hard(By.ID, 'itog').click()

        # Choose table with given subject and class or group
        selectorsubject = Select(self._find_element_try_hard(By.ID, 'sel_pr'))
        self._select_option_try_hard(selectorsubject, subject)
        groupid = 'sel_gr' if groupname.endswith(')') else 'sel_kl'
        selectorgroup = Select(self._find_element_try_hard(By.ID, groupid))
        self._select_option_try_hard(selectorgroup, groupname)

        # Enter grades
        gradeinputs = self._find_elements_try_hard(By.CLASS_NAME, 'sel_mark')
        for gradeinput in gradeinputs:
            inputid = gradeinput.get_property('id')
            studentxpath = '//select[@id="{}"]/parent::*/preceding-sibling'
            studentxpath += '::td[@align="left"]'
            studentxpath = studentxpath.format(inputid)
            studentname = self._find_element_try_hard(
                By.XPATH, studentxpath).get_attribute('textContent')
            studentname = studentname.replace('ё', 'е')
            for studentinfo in info:
                if studentname.startswith(studentinfo[0]):
                    grade = studentinfo[int(inputid[2])+1]
                    if not grade:
                        break
                    samegrade = self._is_same_grade(
                        Select(gradeinput), grade)
                    if not samegrade:
                        if grade == 'Н/А':
                            grade = 'зч'
                            print('{} from {} has N/A'.format(
                                studentinfo[0], classname))
                        self._select_option_try_hard(
                            Select(gradeinput), grade.lower())
                        break

    def get_subjects_and_groups(self):
        # Navigate to "Итоговая ведомость" page
        self.driver.get('https://isko.mosreg.ru/control')
        self._find_element_try_hard(By.ID, 'itog').click()

        # Get all subjects and in which classes and groups it is
        selectorsubject = Select(self._find_element_try_hard(By.ID, 'sel_pr'))
        subjects = [
            option.get_attribute(
                'textContent').strip() for option in selectorsubject.options[1:]]
        subjandgroups = dict.fromkeys(subjects)
        for subject in subjects:
            selectorsubject.select_by_visible_text(subject)
            selectorclass = Select(
                self._find_element_try_hard(By.ID, 'sel_kl'))
            classes = [option.get_attribute(
                'textContent') for option in selectorclass.options[1:]]
            subjandgroups[subject] = classes
            selectorgroup = Select(
                self._find_element_try_hard(By.ID, 'sel_gr'))
            groups = [option.get_attribute(
                'textContent') for option in selectorgroup.options[1:]]
            subjandgroups[subject].extend(groups)

        return subjandgroups

    def snilses(self, classname):
        # Navigate to mosreg page with all students from desired class
        self.driver.get('https://schools.school.mosreg.ru/school.aspx')
        self._find_element_try_hard(By.ID, 'TabClasses').click()
        classlist = self._find_elements_try_hard(
            By.XPATH, '//a[@title="Перейти на страницу класса"]')
        for elem in classlist:
            mosregname = self._mosreg_class_name(classname)
            if elem.get_attribute('textContent') == mosregname:
                elem.click()
                break
        self._find_element_try_hard(By.ID, 'TabMembers').click()

        # Get all snils data
        snilsdata = {}
        editlinks = self._find_elements_try_hard(
            By.XPATH, '//a[@title="Редактировать"]')
        editlinks = [elem.get_attribute('href') for elem in editlinks]
        for link in editlinks:
            self.driver.get(link)
            self._find_element_try_hard(By.ID, 'TabPersonal').click()
            name = self._find_element_try_hard(
                By.ID, 'nlast').get_attribute('value')
            name = name + ' ' + self._find_element_try_hard(
                By.ID, 'nfirst').get_attribute('value')
            name = name + ' ' + self._find_element_try_hard(
                By.ID, 'nmiddle').get_attribute('value')
            name = name.replace('ё', 'е')
            snilsdata[name] = [None, self._find_element_try_hard(
                By.ID, 'personalNumber').get_attribute('value')]

        # Get student ids from ISCO
        self.driver.get('https://isko.mosreg.ru/reestr-stud')
        Select(self._find_element_try_hard(
            By.ID, 'klass_sel')).select_by_visible_text(classname)
        students = self._find_elements_try_hard(
            By.XPATH, '//tr[@class="first filter"]/following-sibling::tr')
        ids = []
        for elem in students:
            style = elem.get_attribute('style')
            if not style.count('display: none;'):
                ids.append(elem.get_attribute('id'))
        for sid in ids:
            name = self._find_element_try_hard(
                By.XPATH,
                '//tr[@id="{}"]/child::td[2]'.format(sid)).get_attribute(
                    'textContent').strip().replace('ё', 'е')
            snilsdata[name][0] = sid

        # Paste snilses to student cards and save them
        script1 = '$(\'<input type="hidden" name="stud_id" value="{}">\')'
        script1 = script1 + '.appendTo(\'#forma_see\');'
        script2 = '$(\'#forma_see\').submit();'
        script3 = '$("#formEdit").append('
        script3 = script3 + '\'<input type="hidden" name="interesId" '
        script3 = script3 + 'value="{}">\');'
        script4 = '$("#formEdit").submit();'
        for name, data in snilsdata.items():
            self.driver.get('https://isko.mosreg.ru/reestr-stud')
            self._find_element_try_hard(By.ID, 'forma_see')
            self.driver.execute_script(script1.format(data[0]))
            self.driver.execute_script(script2)
            dataid = self._find_element_try_hard(
                By.XPATH, '//tr[@data-id]').get_attribute('data-id')
            self.driver.execute_script(script3.format(dataid))
            self.driver.execute_script(script4)
            snils = self._find_element_try_hard(By.ID, 'snils')
            snils.clear()
            snils.send_keys(data[1])
            self._find_element_try_hard(By.ID, 'add').click()

    def login(self, logininfo):
        # Login to ESIA
        self.driver.get('https://esia.gosuslugi.ru/')
        autologin = logininfo.type == Login.BYESIAPASS
        autologin = autologin or logininfo.type == Login.BYBOTHPASSES
        if autologin:
            self._find_element_try_hard(
                By.ID, 'mobileOrEmail').send_keys(logininfo.esianame)
            self._find_element_try_hard(
                By.ID, 'password').send_keys(logininfo.esiapass)
            self._find_element_try_hard(
                By.XPATH,
                '//form[@id="authnFrm"]/descendant::button').click()

        while True:
            try:
                WebDriverWait(self.driver, 5).until(
                    cond.title_contains('Единая система'))
            except TimeoutException:
                continue
            else:
                break
        sleep(2)

        # Login to school portal
        self.driver.get('https://uslugi.mosreg.ru/')
        esiabtn = '//button[@class="g-panel-btn '
        esiabtn += 'g-panel-btn__profile g-panel-btn__default"]'
        self._find_element_try_hard(By.XPATH, esiabtn).click()
        sleep(2)
        self.driver.get('https://sso.mosreg.ru/esiaoauth/login')
        journalbtn = '//button[@class="tile-footer-btn '
        journalbtn += 'school__login-form_button btn-primary"]'
        self.driver.execute_script(
            "arguments[0].click();",
            self._find_element_try_hard(By.XPATH, journalbtn))

        while True:
            try:
                WebDriverWait(self.driver, 5).until(
                    cond.title_contains('Рабочий стол'))
            except TimeoutException:
                continue
            else:
                break

        # Login to ISCO
        if logininfo.type == Login.BYESIAPASS:
            self.driver.get('https://isko.mosreg.ru/shpprof')
            WebDriverWait(self.driver, 10).until(cond.title_is(
                'Система оценки качества образования'))
        elif logininfo.type == Login.BYBOTHPASSES:
            self.driver.get('https://isko.mosreg.ru/')
            self._find_element_try_hard(
                By.NAME, 'log').send_keys(logininfo.isconame)
            self._find_element_try_hard(
                By.NAME, 'pas').send_keys(logininfo.iscopass)
            self._find_element_try_hard(By.NAME, 'save').click()
        else:
            self.driver.get('https://isko.mosreg.ru/')

        while True:
            try:
                WebDriverWait(self.driver, 5).until(
                    cond.title_contains(
                        'Система оценки качества образования'))
            except TimeoutException:
                continue
            else:
                break

    def exit(self):
        sleep(2)
        self.driver.quit()

    def _mosreg_class_name(self, isconame):
        return isconame[:-1] + '-' + isconame[-1].lower()

    def _find_element_try_hard(self, by, value):
        element = None
        while True:
            try:
                element = self.driver.find_element(by, value)
            except cond.NoSuchElementException:
                continue
            else:
                return element

    def _find_elements_try_hard(self, by, value):
        elements = None
        while True:
            elements = self.driver.find_elements(by, value)
            if elements == []:
                continue
            else:
                return elements

    def _select_option_try_hard(self, select, value):
        while True:
            try:
                select.select_by_visible_text(value)
            except cond.NoSuchElementException:
                continue
            else:
                return

    def _is_same_grade(self, select, grade):
        return select.first_selected_option.get_attribute(
            'textContent') == grade


def mainitog():
    paster = MosregToISCO()
    logininfo = Login('logindata.yml', Login.BYBOTHPASSES)
    paster.login(logininfo)
    subjandgroups = paster.get_subjects_and_groups()

    # Remove baseschool
    sng = copy.deepcopy(subjandgroups)
    for subject, groups in subjandgroups.items():
        for groupname in groups:
            classname = re.search('[\d]{1,2}[А-Е]', groupname).group(0)
            baseschool = classname.startswith('1') or classname.startswith('2')
            baseschool = baseschool or classname.startswith('3')
            baseschool = baseschool or classname.startswith('4')
            baseschool = baseschool and len(classname) == 2
            highschool = classname.startswith('1') and len(classname) == 3
            if baseschool or highschool:
                sng[subject].remove(groupname)
        if not sng[subject]:
            sng.pop(subject)

    # Remove all 7-9 classes from mathematics
    subjandgroups = copy.deepcopy(sng)
    for groupname in subjandgroups.get('математика'):
        classname = re.search('[\d]{1,2}[А-Е]', groupname).group(0)
        toobigformath = classname.startswith('7') or classname.startswith('8')
        toobigformath = toobigformath or classname.startswith('9')
        if toobigformath:
            sng['математика'].remove(groupname)

    # Remove all except needed
    """ for subject in subjandgroups.keys():
        exclude = subject == 'алгебра' 
        exclude = exclude or subject == 'алгебра и начала анализа'
        exclude = exclude or subject == 'английский язык'
        exclude = exclude or subject == 'астрономия'
        exclude = exclude or subject == 'биология'
        exclude = exclude or subject == 'всеобщая история'
        exclude = exclude or subject == 'география'
        exclude = exclude or subject == 'геометрия'
        exclude = exclude or subject == 'информатика'
        exclude = exclude or subject == 'искусство (мхк)'
        exclude = exclude or subject == 'история'
        exclude = exclude or subject == 'история России'
        exclude = exclude or subject == 'литература'
        exclude = exclude or subject == 'математика'
        exclude = exclude or subject == 'музыка'
        exclude = exclude or subject == 'немецкий язык'
        exclude = exclude or subject == 'обществознание'
        exclude = exclude or subject == 'основы безопасности жизнедеятельности'
        exclude = exclude or subject == 'право'
        exclude = exclude or subject == 'русский язык'
        exclude = exclude or subject == 'технология'
        exclude = exclude or subject == 'физика'
        exclude = exclude or subject == 'физическая культура'
        exclude = exclude or subject == 'французский язык'
        if exclude:
            sng.pop(subject)
    for groupname in subjandgroups['химия']:
        classname = re.search('[\d]{1,2}[А-Е]', groupname).group(0)
        if classname.startswith('10') or classname.startswith('10'):
            sng['химия'].remove(groupname) """

    # Copy grades
    paster.exit()
    for subject, classes in sng.items():
        print('======  ' + subject + '  ======')
        paster = MosregToISCO()
        paster.login(logininfo)
        for classname in classes:
            paster.itog(subject, classname)
        paster.exit()


def mainsnilses():
    paster = MosregToISCO()
    logininfo = Login('logindata.yml', Login.BYBOTHPASSES)
    paster.login(logininfo)
    paster.snilses('5В')


if __name__ == '__main__':
    mainitog()
