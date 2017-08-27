/*  
    This sample demonstrates how to use the Data-Driven Testing plug-in 
    and how to create data-driven tests.
    The application form is populated several times with data extracted from the 
    <TestComplete Samples>\Common\Data-Driven Testing\TestBook.xlsx Excel file.
    
    The script loads the Orders.exe application from the
    <TestComplete Samples>\Desktop\Orders\C#\bin\Release\ folder.
   
    Requirements:
      Microsoft Excel 2007 must be installed on your computer.
      Data-Driven Testing plug-in must be installed in TestComplete.    
*/

function OpenForm(mainForm)
{
  mainForm.MainMenu.Click("Orders|New order...");
}

function PopulateForm(groupBox, Driver)
{
  var numericUpDown, textBox; 

  groupBox.ProductNames.ClickItem(Driver.Value(1));
  numericUpDown = groupBox.Quantity;
  numericUpDown.Click(8, 7);
  numericUpDown.wValue = Driver.Value(2);
  groupBox.Date.wDate = Driver.Value(3);
  textBox = groupBox.Customer;
  textBox.Click(69, 12);
  textBox.wText = Driver.Value(0);
  textBox = groupBox.Street;
  textBox.Click(4, 8);
  textBox.wText = Driver.Value(4);
  textBox = groupBox.City;
  textBox.Click(12, 14);
  textBox.wText = Driver.Value(5);
  textBox = groupBox.State;
  textBox.Click(21, 6);
  textBox.wText = Driver.Value(6);
  textBox = groupBox.Zip;
  textBox.Click(14, 14);
  textBox.wText = Driver.Value(7);
  groupBox.WinFormsObject(Driver.Value(8)).ClickButton();
  textBox = groupBox.CardNo;
  textBox.Click(35, 15);
  textBox.wText = Driver.Value(9);
  groupBox.ExpDate.wDate = Driver.Value(10);
}

function Checkpoint(Driver)
{
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.Price, "wText", cmpEqual, Driver.Value(11), false);
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.Discount, "wText", cmpEqual,Driver.Value(12), false);
  aqObject.CheckProperty(Aliases.Orders.OrderForm.Group.groupBox1.Total, "wText", cmpEqual, Driver.Value(13), false);
}

function CloseForm(orderForm)
{
  orderForm.ButtonOK.ClickButton();
}

function CloseApplication(mainForm, orders)
{
  mainForm.Close();
  orders.dlgConfirmation.btnNo.ClickButton();
}

function Main()
{
  var orders, mainForm, orderForm, groupBox, Driver;

  Driver = DDT.ExcelDriver("../../TestBook.xlsx", "TestSheet", true);
  TestedApps.RunAll();
  orders = Aliases.Orders;
  mainForm = orders.MainForm;
  while(!Driver.EOF())
  {
    OpenForm(mainForm);
    orderForm = orders.OrderForm;
    groupBox = orderForm.Group;
    PopulateForm(groupBox, Driver);
    Checkpoint(Driver);
    CloseForm(orderForm);
    Driver.Next();
  }
  CloseApplication(mainForm, orders);
}

