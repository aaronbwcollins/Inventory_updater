import easygui as g
import gspread

# Version 1.1

### Scope of creds
scope = ['https://spreadsheets.google.com/feeds']

### File path of Service Account creds
file_path = "filepath\\to\\service_account.json"

### Service account authentication
gc = gspread.service_account(filename=file_path)
### 
wks = gc.open("Inventory_1") ##Changed for Dev enviorment
### Creating variable to work from of the worksheet
worksheet = wks.worksheet("Sheet1")

def add():
### Add item function loop
    message = ""
    while True:
        scan = barcode_scan("Please Scan Barcode to add\n"+str(message))
        if scan == "Quit":
            break
        else:
            message = add_item(scan)


def add_item(scan):
### Function to add multiple quanties of give item
    try:
        qty_to_update = g.enterbox("What Quantity to Add?")
        find_cell = find_barcode(scan)
        cellRow = int(find_cell.row)
        cellCol = int(find_cell.col)
        values_list = worksheet.row_values(cellRow)
        cellCol = cellCol+2
        current_qty = int(values_list[2]) + int(qty_to_update)
        worksheet.update_cell(cellRow, cellCol, current_qty)
        message = ("%s has been increased to %s" % (values_list[1],current_qty))
        return message
    except:
        message = ""
        if g.ccbox("Unknown Barcode\nWould your Like to Add Product?"):
            message = new_product_1()
            return message
        else:
            return message

def audit_inventory():
### Function to audit inventory of warehouse
    for i in range(2,len(worksheet.get_all_values())+1):
        values_list = worksheet.row_values(i)
        if values_list[0] != '':
            barcode = values_list[0]
            cell = worksheet.find(barcode)
            cellRow = int(cell.row)
            cellCol = int(cell.col)
            cellCol = cellCol + 2
            current_item = str(values_list[1])
            reported_stock = str(values_list[2])
            qty_to_update = audit(current_item, reported_stock)
            print(qty_to_update)
            worksheet.update_cell(cellRow,cellCol,qty_to_update)
            confirmation_audit(current_item, str(int(reported_stock) - int(qty_to_update)), str(qty_to_update))
    audit_complete()

def audit(item, qty):
### Allows user to enter current stock during an audit
    audit_qty = g.enterbox("Audit Mode\nReported Stock of %s is %s\nPlease Enter Current Stock" % (item, qty))
    return(audit_qty)

def confirmation_audit(item, deviation, qty):
### Prints message when deviation is found
    if int(deviation) == 0:
        pass
    else:
        g.msgbox("Stock deviation for %s is %s\nStock now set at %s" % (item, deviation, qty))

def audit_complete():
    g.msgbox("Audit Mode Complete!")

def new_product():
### Adds new product but askes first if you would like to addd one
    while True:
        if g.ccbox("Would you like to add a new product?"):
            list_of_lists = worksheet.get_all_values()
            current_number = len(list_of_lists)
            current_number = str(current_number+1)    
            barcode = g.enterbox("Please Scan Barcode to add")
            worksheet.update_acell("A%s" %(current_number),barcode)
            item = g.enterbox("Please Type Name of Product")
            worksheet.update_acell("B%s" %(current_number),item)
            quantity = g.enterbox("Please Type Quantity")
            worksheet.update_acell("C%s" %(current_number),quantity)
        else:
            break


def new_product_1():
### Adds new product but doesn't ask if you need to
    list_of_lists = worksheet.get_all_values()
    current_number = len(list_of_lists)
    current_number = str(current_number+1)    
    barcode = g.enterbox("Please Scan Barcode to add")
    worksheet.update_acell("A%s" %(current_number),barcode)
    item = g.enterbox("Please Type Name of Product")
    worksheet.update_acell("B%s" %(current_number),item)
    quantity = g.enterbox("Please Type Quantity")
    worksheet.update_acell("C%s" %(current_number),quantity)
    message = "%s added to worksheet" % (item)
    return message

def barcode_scan(message):
### Simple fuction to create input box for barcode
    scan = g.enterbox(msg=message, title="Inventory Updater")
    return scan

def intial_barcode_scan():
### Intial scan to select mode
    mode = g.enterbox(msg="Please Select Mode", title="Inventory Updater")
    return mode

def reduction(scan):
### Will find and reduce value of cell by 1 then return message of what item has been reduced
    try:
        qty_to_update = -(int('1'))
        find_cell = find_barcode(scan)
        cellRow = int(find_cell.row)
        cellCol = int(find_cell.col)
        values_list = worksheet.row_values(cellRow)
        cellCol = cellCol+2
        current_qty = int(values_list[2]) + qty_to_update
        worksheet.update_cell(cellRow, cellCol, current_qty)
        message = ("%s has been reduced to %s" % (values_list[1],current_qty))
        #g.msgbox(message)
        return message
    except:
        message = ""
        if g.ccbox("Unknown Barcode\nWould your Like to Add Product?"):
            message = new_product_1()
            return message
        else:
            return message


def find_barcode(barcode):
### Will find barcode and grab qty cell
    current_cell = worksheet.find(barcode)
    return current_cell


def quick_pick():
### Runs loop for reducting items    
    message = ""
    while True:
        scan = barcode_scan("Please Scan Barcode to reduce\n"+str(message))
        if scan == "Quit":
            break
        elif scan == "Convert":
            convert()
        else:
            message = reduction(scan)

def convert():
    message = ""
    scan = barcode_scan("Convert Mode\nPlease Scan Barcode to reduce")
    reduction(scan)
    add()



def main_loop():
### Main loop for progam    
    while True:
        mode = intial_barcode_scan()
        #print(scan)
        if mode == 'Add':
            add()
        elif mode == 'Quickpick':
            quick_pick()
        elif mode == 'Audit':
            audit_inventory()
        elif mode == "New Product":
            new_product()
        elif mode == 'Quit':
            break
        elif mode == 'Help':
            g.msgbox('Modes are\nAdd, Audit, New, Quickpick, or Quit')
        elif mode == 'Convert':
            convert()
        else:
            break



if __name__ == "__main__":
    main_loop()


