import urllib.request
import re
from bs4 import BeautifulSoup as Bsoup
import xlsxwriter

# root URL to be attached to found links
rootUrl = "http://polymerdatabase.com/"
# root URL to be attached to found links
root2Url = "http://polymerdatabase.com/polymer%20index/"
# main page URL that we'll use to access the other pages
hostUrl = "http://polymerdatabase.com/home.html"
# links to all polymers found
polymer_links = []

headings = ["NAME","FORMULA","PROPERTY: Molecular Weight of Repeat Unit (g/mol)",
            "PROPERTY: Van-der-waals Volume (mL/mol)","PROPERTY: Molar Volume (mL/mol)"
            ,"PROPERTY: Density (g/mL)","PROPERTY: Solubility Parameter (MPa^1/2)",
            "PROPERTY: Molar Cohesive Energy (J/mol)","PROPERTY: Tg (K)",
            "PROPERTY: Cp (J/(mol*K))","PROPERTY: Entanglement Molecular Weight (g/mol)",
            "PROPERTY: Index of Rrefraction (n)"]
theData = [headings]

# Creates the tag soup
def make_soup(url):
    # query the website and return the html to the variable 'page'
    directory = urllib.request.urlopen(url)
    # parse the html using beautifulsoup and store in variable 'html'
    soup = Bsoup(directory, 'html.parser')
    return soup

# ___Functions___

# function to find polymer classes from a table
def in_table_dir(tag):
    try:
        if (tag.contents[0]['href'] == '#' or tag.contents[0]['href'] == '#.html'):
            return False
    except:
        return False
    return (tag.parent.name == 'td') and (tag.name == 'p')

# function to find polymers from list
def in_list(href):
    return href and re.compile("polymers").search(href) and re.compile(".html").search(href) and not re.compile("polymer classes").search(href)

# function to find the name of the polymer
def data_in_table(tag):
    #return (tag.name == 'tr') and ((tag.contents[0].string == "Molecular Weight of Repeat unit") or (tag.contents[0].string == "Glass Transition Temperature "))
    return (tag.parent.name == 'div') and (tag.name == 'b')

# function to find the chemical structure from a table
def find_smiles(tag):
    try:
        if (tag.contents[0] == 'SMILES'):
            return True
    except:
        return False

# ---Functions---

# Branch from the home page to find the secondary directories
def branch_from_home():
    # store URL's for each directory (A-Z)
    AtoZ = []

    # home url same as A-B directory
    AtoZ.append(hostUrl)

    soup = make_soup(hostUrl)
    for tag in soup.find_all('li'):
        if (tag.text == 'C - D' or tag.text == 'E - F' or
                    tag.text == 'G - L' or tag.text == 'M - P' or tag.text == 'R - Z'):
            url = rootUrl + tag.contents[0]['href']
            url = url.replace(" ", "%20")
            AtoZ.append(url)

    i = 0
    for link in AtoZ:
        if (i >= 0):
            #print("___Directory:\n" + link + "\n___Polymer Types:")
            branch_from_Directory(link)
        i += 1
    populate_dataSet()

def branch_from_Directory(page):
    # store URL's for each Polymer class
    polymers = []
    soup = make_soup(page)
    for tag in soup.find_all(in_table_dir):
        if (tag.contents[0]['href'].__contains__("index")):
            url = rootUrl + tag.contents[0]['href']
        else:
            url = root2Url + tag.contents[0]['href']
        url = url.replace(" ", "%20")
        polymers.append(url)
    i = 0
    for link in polymers:
        if (i >= 0):
            branch_from_url(link)
        i += 1

def branch_from_url(url):
    soup = make_soup(url)

    i = 0
    for tag in soup.find_all(href=in_list):
        if (i >= 0):
            poly = rootUrl + tag['href']
            poly = poly.replace(" ", "%20")
            poly = poly.replace("/../", "/")
            polymer_links.append(poly)
        i += 1

# Table parsing code
def get_data(url):
    soup = make_soup(url)
    # find all tables in the page
    tables = soup.findAll("table")
    properties = []
    poly_name = soup.find_all(data_in_table)[0].string
    properties.append(poly_name)
    once = True

    tableNum = len(tables) - 1
    try:
        #lines = tables[1].findAll("tr")
        smiles = soup.find_all(find_smiles)[0].parent
        formula = smiles.text.split(" ")[1]
        properties.append(formula)
    except:
        print("___Failed___")

    # for each row in the first table
    for player in tables[tableNum].findAll("tr"):

        # parse over each column in the row, store in values
        values = player.findAll("td")

        if (once):
            once = False
            continue

        # can now access the elements of the rows by element ([0], [1], [2], etc) to access the information needed
        try:
            if (values[3].contents[0] == None):
                val = values[2].contents[0]
            else:
                val = values[3].contents[0]
            l = []
            for t in val.split():
                try:
                    l.append(float(t))
                except ValueError:
                    pass
            properties.append(l[0])
        except:
            try:
                val = values[2].contents[0]
                l = []
                for t in val.split():
                    try:
                        l.append(float(t))
                    except ValueError:
                        pass
                properties.append(l[0])
            except:
                properties.append(" ")
                return None
    return properties

def populate_dataSet():
    global polymer_links
    global theData

    # Create Dataset within a double array in python
    for link in polymer_links:
        data = get_data(link)
        if (not (data == None)):
            theData.append(data)
        print("Polymer: " + link)
    print("# Polymers: %d" % len(polymer_links))

    # Export the dataset into an excel file
    write_to_excel(theData)

def write_to_excel(data):
    # Create a workbook and add a worksheet.
    workbook = xlsxwriter.Workbook('PolymerDataSheet05.xlsx')
    worksheet = workbook.add_worksheet()

    # Start from the first cell. Rows and columns are zero indexed.
    row = 0
    col = 0

    print (data[0])
    # Iterate over the data and write it out row by row.
    for line in (data):
        for prp in line:
            worksheet.write(row, col, prp)
            col += 1
        row += 1
        col = 0

    workbook.close()


#____________________________________________________________________________________

# Create a populated dataset
branch_from_home()