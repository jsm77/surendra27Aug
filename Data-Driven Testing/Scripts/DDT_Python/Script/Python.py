#    This sample demonstrates how to use the Data-Driven Testing plug-in 
#    and how to create data-driven tests.
#    The application form is populated several times with data extracted from the 
#    <TestComplete Samples>\Common\Data-Driven Testing\TestBook.xlsx Excel file.
#    
#    The script loads the Orders.exe application from the
#    <TestComplete Samples>\Desktop\Orders\C#\bin\Release\ folder.
#   
#    Requirements:
#      Microsoft Excel 2007 must be installed on your computer.
#      Data-Driven Testing plug-in must be installed in TestComplete.    


def openForm(mainForm):
  mainForm.MainMenu.Click("Orders|New order...")

def populateForm(groupBox, driver):
  groupBox.ProductNames.ClickItem(driver.Value[1])
  numericUpDown = groupBox.Quantity
  numericUpDown.Click(8, 7)
  numericUpDown.wValue = driver.Value[2]
  groupBox.Date.wDate = driver.Value[3]
  textBox = groupBox.Customer
  textBox.Click(69, 12)
  textBox.wText = driver.Value[0]
  textBox = groupBox.Street
  textBox.Click(4, 8)
  textBox.wText = driver.Value[4]
  textBox = groupBox.City
  textBox.Click(12, 14)
  textBox.wText = driver.Value[5]
  textBox = groupBox.State
  textBox.Click(21, 6)
  textBox.wText = driver.Value[6]
  textBox = groupBox.Zip
  textBox.Click(14, 14)
  textBox.wText = driver.Value[7]
  groupBox.WinFormsObject(driver.Value[8]).ClickButton()
  textBox = groupBox.CardNo
  textBox.Click(35, 15)
  textBox.wText = driver.Value[9]
  groupBox.ExpDate.wDate = driver.Value[10]

def checkpoint(driver):
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.Price, "wText", cmpEqual, driver.Value[11], False)
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.Discount, "wText", cmpEqual, driver.Value[12], False)
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.groupBox1.Total, "wText", cmpEqual, driver.Value[13], False)

def closeForm(orderForm):
  orderForm.ButtonOK.ClickButton()

def closeApplication(mainForm, orders):
  mainForm.Close()
  orders.dlgConfirmation.btnNo.ClickButton()

def main():
  driver = DDT.ExcelDriver("../../TestBook.xlsx", "TestSheet", True)
  TestedApps.RunAll()
  orders = Aliases.Orders
  mainForm = orders.MainForm
  while not driver.EOF():
    openForm(mainForm)
    orderForm = orders.OrderForm
    groupBox = orderForm.Group
    populateForm(groupBox, driver)
    checkpoint(driver)
    closeForm(orderForm)
    driver.Next()

  closeApplication(mainForm, orders)
