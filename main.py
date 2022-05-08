#
# Author: yilmazturkm
# Author Github: https://github.com/yilmazturkm
# Author Url: https://yilmazturk.gen.tr
#
# Product details scrapping application for trendyol online store
#


import sys
import requests
import pathlib
import xlsxwriter
from selenium import webdriver
from bs4 import BeautifulSoup
from PyQt5.QtGui import QIcon, QCursor
from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal
import time
from PyQt5.QtWidgets import (
    QWidget, QTextEdit, QApplication,
    QLabel, QTextEdit,
    QLineEdit, QGroupBox,
    QVBoxLayout, QHBoxLayout,
    QFormLayout, QPushButton,
    QComboBox, QScrollBar)

class Worker(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(str)
    def __init__(self, link, count, page, fileName):
        super().__init__()
        self.link = link
        self.count = count
        self.page = page
        self.fileName = fileName

    def run(self):
        warningText = ""
        if len(self.page) < 1:
            warningText += "- Choose the page you want to scrap!\n"
        try:
            if int(self.count) < 1:
                warningText += "- Number of product can not be empty or 0!\n"
            else:
                self.count = int(self.count)
        except Exception as e:
            warningText += "- Only integer for number of product field!\n"
            print(e)
        if len(self.fileName) < 0:
            warningText += "- File name can not be empty!"
        else:
            fileName = self.fileName + ".xlsx"
            file = pathlib.Path(fileName)
            if file.exists():
                warningText = "- File name exists. Type another file name!"
        
        if len(warningText) > 0:
            self.progress.emit(warningText)
        else:
            productLinks = self.getProductLinks(self.link, self.count, self.page)
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            browser = webdriver.Chrome(options=options)
            browser.get("https://www.trendyol.com")
            row = 1
            fileName = self.fileName + ".xlsx"
            workbook = xlsxwriter.Workbook(fileName)
            worksheet = workbook.add_worksheet()
            worksheet.write("A1", "Product Link")
            worksheet.write("B1", "Product Name")
            worksheet.write("C1", "Price")
            worksheet.write("D1", "Sale Price")
            worksheet.write("E1", "Cart Price")
            worksheet.write("F1", "Product Variants")
            worksheet.write("G1", "Sizes")
            worksheet.write("H1", "Images")
            worksheet.write("I1", "Seller")
            worksheet.write("J1", "Description")
            for i in productLinks:
                self.progress.emit(f"{productLinks.index(i) + 1} - Getting product details... {i}")
                product = self.getProductDetails(i, browser)
                values = list(product.values())
                for j in values:
                    worksheet.write(row, values.index(j), j)
                row += 1
            workbook.close()
            browser.quit()
            self.progress.emit("Details of all products saved.")
        self.finished.emit()
    
    def getProductLinks(self, archiveLink, productCount, scrapFrom):
        pn = 1
        x = 1
        exit = 0
        linkList = []
        message = ""
        maxCount = int(productCount/24) + 1
        while exit == 0:
            url = ""
            if scrapFrom == "Category":
                url = archiveLink + "?pi=" + str(pn)
            elif scrapFrom == "Seller":
                url = archiveLink + "&pi=" + str(pn)
            if len(url) > 0:
                r = requests.get(url).text
                content = BeautifulSoup(r, "lxml")
                products = content.find_all("div", {"class": "p-card-wrppr"})
                for i in products:
                    link = i.find_all("a", href=True)[0]['href']
                    self.progress.emit(f"{x} - {link}")
                    linkList.append(link)
                    if x == productCount:
                        self.progress.emit("\nProduct links saved. Now product details will be saved!\n")
                        exit = 1
                        break
                    x += 1
                pn += 1
        return linkList

    def getProductDetails(self, productLink, browser):
        productDetails = {}
        url = "https://www.trendyol.com" + productLink
        browser.get(url)
        time.sleep(2)
        productDetails["productUrl"] = url
        try:
            productTitle = browser.find_element_by_class_name("pr-new-br").text
            productDetails["productTitle"] = productTitle
        except Exception as e:
            productDetails["productTitle"] = ""
            self.progress.emit("- Could not get product title!\n")
        try:
            productOrgPrice = browser.find_element_by_class_name("prc-org").text
            productOrgPrice = productOrgPrice.replace(" TL", "")
            productDetails["productOrginalPrice"] = productOrgPrice
        except Exception as e:
            productDetails["productOrginalPrice"] = ""
            self.progress.emit("- Could not find product price\n")
        try:
            productDiscountPrice = browser.find_element_by_class_name("prc-slg").text
            productDiscountPrice = productDiscountPrice.replace(" TL", "")
            productDetails["productDiscountPrice"] = productDiscountPrice
        except Exception as e:
            productDetails["productDiscountPrice"] = ""
            self.progress.emit("- Could not find sale price\n")
        try:
            productCartPrice = browser.find_element_by_class_name("prc-dsc").text
            productCartPrice = productCartPrice.replace(" TL", "")
            productDetails["productCartPrice"] = productCartPrice
        except Exception as e:
            productDetails["productCartPrice"] = ""
            self.progress.emit("- Could not find cart price\n")
        try:
            productVer = browser.find_element_by_xpath("/html/body/div[1]/div[5]/main/div/div[2]/div[1]/div[2]/div[2]/section/div[2]/div")
            productVers = productVer.find_elements_by_tag_name("a")
            productVersions = []
            for i in productVers:
                productVersions.append(i.get_attribute("title"))
            productVersions = ", ".join(map(str, productVersions))
            productDetails["productVersions"] = productVersions
        except Exception as e:
            productDetails["productVersions"] = ""
            self.progress.emit("- Could not find product variants\n")
        try:
            productSize = browser.find_elements_by_class_name("sp-itm")
            productSizes = []
            for i in productSize:
                productSizes.append(i.text)
            productSizes = ", ".join(map(str, productSizes))
            productDetails["productSizes"] = productSizes
        except Exception as e:
            productDetails["productSizes"] = ""
            self.progress.emit("- Could not find product sizes\n")
        try:
            productImage = browser.find_element_by_xpath("/html/body/div[1]/div[5]/main/div/div[2]/div[1]/div[1]/div/div[1]/div/div")
            productImgs = productImage.find_elements_by_tag_name("img")
            productImages = []
            for i in productImgs:
                productImages.append(i.get_attribute("src"))
            productImages = " ".join(map(str, productImages))
            productDetails["productImages"] = productImages
        except Exception as e:
            productDetails["productImages"] = ""
            self.progress.emit("- Could not find product image\n")
        try:
            productSeller = browser.find_element_by_class_name("merchant-text").text
            productDetails["productSeller"] = productSeller
        except Exception as e:
            productDetails["productSeller"] = ""
            self.progress.emit("- Could not find product seller\n")
        try:
            productInfo = browser.find_element_by_xpath("/html/body/div[1]/div[5]/main/div/section/div/div").text
            productInfo = productInfo.replace("\n", " ")
            productDetails["productInfo"] = productInfo
        except Exception as e:
            productDetails["productInfo"] = ""
            self.progress.emit("- Could not find product description\n")
        
        return productDetails


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.setStyleSheet(open('mystylesheet.css').read())

    def initUI(self):
        mh = 600
        mw = 1000
        lh = 24
        lw = 220
        linkLabel = QLabel("Category or Seller page url", self)
        linkLabel.setFixedHeight(lh)
        linkLabel.setFixedWidth(lw)
        sourceLabel = QLabel("Which page will be scrapped", self)
        sourceLabel.setFixedHeight(lh)
        numberLabel = QLabel("Number of Product", self)
        numberLabel.setFixedHeight(lh)
        fileNameLabel = QLabel("File name", self)
        fileNameLabel.setFixedHeight(lh)
        buttonLabel = QLabel("", self)
        buttonLabel.setFixedHeight(lh)
        self.linkField = QLineEdit("https://...")
        self.linkField.setFixedWidth(250)
        self.sourceField = QComboBox()
        self.sourceField.addItems(["","Category", "Seller"])
        self.sourceField.setMaximumWidth(100)
        self.numberField = QLineEdit()
        self.numberField.setFixedWidth(50)
        self.fileNameField = QLineEdit()
        self.fileNameField.setFixedWidth(140)
        self.button = QPushButton("Search Products")
        self.button.setFixedWidth(100)
        self.button.setFixedHeight(lh)
        self.button.setCursor(QCursor(Qt.PointingHandCursor))
        self.button.clicked.connect(self.runLongTask)
        textAreaBox = QGroupBox("Enter the information above and click \"Search Products\" button")
        scrollBar = QScrollBar()
        self.textArea = QTextEdit()
        self.textArea.setReadOnly(True)
        self.textArea.setVerticalScrollBar(scrollBar)
        textAreaLayout = QHBoxLayout()
        textAreaLayout.addWidget(self.textArea)
        textAreaBox.setLayout(textAreaLayout)
        topLayout=QFormLayout()
        topLayout.addRow(linkLabel, self.linkField)
        topLayout.addRow(sourceLabel, self.sourceField)
        topLayout.addRow(numberLabel, self.numberField)
        topLayout.addRow(fileNameLabel, self.fileNameField)
        topLayout.addRow(buttonLabel,self.button)
        topLayout.addRow(textAreaBox)      
        mainLayout = QVBoxLayout()
        mainLayout.addLayout(topLayout)
        self.setFixedSize(mw, mh)
        self.setLayout(mainLayout)
        self.setWindowTitle('Save Products Info')
        self.setWindowIcon(QIcon('web.png'))
        self.show()
    
    def reportProgress(self, n):
        self.textArea.append(n)

    def runLongTask(self):
        link = self.linkField.text()
        count = self.numberField.text()
        page = self.sourceField.currentText()
        fileName = self.fileNameField.text()
        self.textArea.setText("")
        # Step 2: Create a QThread object
        self.thread = QThread()
        # Step 3: Create a worker object
        self.worker = Worker(link, count, page, fileName)
        # Step 4: Move worker to the thread
        self.worker.moveToThread(self.thread)
        # Step 5: Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.reportProgress)
        # Step 6: Start the thread
        self.thread.start()

        # Final resets
        self.button.setEnabled(False)
        self.thread.finished.connect(
            lambda: self.button.setEnabled(True)
        )
        """
        self.thread.finished.connect(
            lambda: self.textArea.setText("İşlem tamamlandı!")
        )
        """


def main():

    app = QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()