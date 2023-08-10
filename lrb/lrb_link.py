import time

import openpyxl
from selenium.common import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from lrb.common import util
from lrb.common.Model import Lrb, Excel


class Item:
    row = 0
    id = ''
    link = ''
    log = ''

    def __init__(self, _row, _id, _link):
        self.row = _row
        self.id = _id
        self.link = _link

    def to_string(self):
        return 'row=%s, id=%s' % (self.row, self.id)


class LrbLink(Lrb):

    def __read_excel(self):
        xlsx = openpyxl.load_workbook(self._excel_path)
        active = xlsx.active
        max_row = active.max_row
        max_column = active.max_column
        self._emit('读取excel最大行数：{}，最大列数：{}', max_row, max_column)

        items = []
        for j in range(1, max_row + 1):
            if active.cell(row=j, column=3).value == 'success':
                continue
            item = Item(j, active.cell(row=j, column=1).value, active.cell(row=j, column=2).value)
            items.append(item)
        self._emit('共读取待更换监测链接记录 {}条', len(items))
        self.excel = Excel(self._excel_path, xlsx, active, max_row, max_column, items)

    def __change_link(self):
        manage = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="manage-list"), self.edge)
        id_input = util.wait_for_find_ele(
            lambda d: manage.find_element(by=By.TAG_NAME, value="input"), self.edge)
        id_input.send_keys('')
        id_input.clear()
        id_input.send_keys(self.item.id)
        id_input.send_keys(Keys.ENTER)
        time.sleep(1)

        edit_retry = 3
        while edit_retry > 0:
            try:
                edit_div = util.wait_for_find_ele(
                    lambda d: d.find_element(by=By.CLASS_NAME, value="css-ece9u5"), self.edge)
                edit_a = util.wait_for_find_ele(
                    lambda d: edit_div.find_elements(by=By.TAG_NAME, value="a"), self.edge)
                edit_a[0].click()
                break
            except StaleElementReferenceException:
                edit_retry -= 1
                time.sleep(0.5)

        self.edge.switch_to.window(self.edge.window_handles[-1])

        table = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="el-table__fixed-body-wrapper"), self.edge)
        td = util.wait_for_find_ele(
            lambda d: table.find_elements(by=By.TAG_NAME, value="td"), self.edge)
        change_btn = util.wait_for_find_ele(
            lambda d: td[13].find_element(by=By.CLASS_NAME, value="link-text"), self.edge)
        change_btn.click()

        clear_btn = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-1jjt3ne"), self.edge)
        clear_btn.click()
        link_input = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-968ze5"), self.edge)
        link_input.clear()
        link_input.send_keys(self.item.link)

        save_btn = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-php29w"), self.edge)
        save_btn.click()
        time.sleep(0.7)

        finish_btn = util.wait_for_find_ele(
            lambda d: d.find_element(by=By.CLASS_NAME, value="css-r7neow"), self.edge)
        finish_btn.click()
        time.sleep(1)
        self.edge.close()
        self.edge.switch_to.window(self.edge.window_handles[0])

    def begin(self):
        super().execute(
            '监测链接修改',
            self.__read_excel,
            util.creative_page,
            self.__change_link,
            3
        )


def main():
    filepath = r'C:\Users\mages\Desktop\创意id+监测链接.xlsx'
    username = ''
    password = ''
    ll = LrbLink(filepath, username, password)
    # noinspection PyUnresolvedReferences
    ll.msg.connect(lambda m: print(m))
    ll.begin()


if __name__ == '__main__':
    main()
