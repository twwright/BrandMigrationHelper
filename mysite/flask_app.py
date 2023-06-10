from flask import Flask, request, Response
import services
import products
import series
import customer_series
import employees

app = Flask(__name__)

# METHODS

def transform_services():
    with open('services.xlsx', 'rb') as f:
        return services.services(f)

def transform_products():
    with open('products.xlsx', 'rb') as f:
        return products.products(f)

def transform_inventory():
    with open('inventory.xlsx', 'rb') as f:
        return products.inventory(f)

def transform_series():
    with open('series.xlsx', 'rb') as f:
        return series.series(f)

def transform_customer_series():
    with open('cseries.xlsx', 'rb') as f:
        return customer_series.cseries(f)

def transform_employees():
    with open('employees.xlsx', 'rb') as f:
        return employees.employees(f)

# PRIMARY GET ROUTE

@app.route('/')
def form():
  return """
    <html>
        <head>
            <title>Brand Data Helpers</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/kognise/water.css@latest/dist/dark.min.css">
            <style>
            .tooltip {
                position: relative;
                display: inline-block;
            }

            .tooltip .tooltiptext {
              visibility: hidden;
              width: 350px;
              font-size: 14px;
              background-color: black;
              color: #fff;
              text-align: center;
              padding: 5px 5px;
              border-radius: 6px;
              position: absolute;
              z-index: 1;
              top: -5px;
              left: 110%;
            }

            .tooltip .tooltiptext::after {
              content: "";
              position: absolute;
              top: 50%;
              right: 100%;
              margin-top: 10px;
              border-width: 5px;
              border-style: solid;
              border-color: transparent black transparent transparent;
            }

            .tooltip:hover .tooltiptext {
              visibility: visible;
            }
            </style
        </head>
        <body>
            <h1>Data Conversion Tools</h1>

        <div id="options">
            <h2>Brand Services</h2>

            <form action="/transform_services" method="post" enctype="multipart/form-data">
                <input type="file" name="services_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Services as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "services-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file using the Brand Uploader.</li>
                    </ul>
                </p>
            </details>


        <hr>

            <h2>Products</h2>

            <form action="/transform_products" method="post" enctype="multipart/form-data">
                <input type="file" name="products_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>

            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Services as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "services-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file using the Brand Uploader.</li>
                    </ul>
                </p>
            </details>

        <hr>

            <h2>Product Inventory</h2>

            <form action="/transform_inventory" method="post" enctype="multipart/form-data">
                <input type="file" name="inventory_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>

            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Services as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "services-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file using the Brand Uploader.</li>
                    </ul>
                </p>
            </details>

        <hr>

            <h2>Brand Series</h2>

            <form action="/transform_series" method="post" enctype="multipart/form-data">
                <input type="file" name="series_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>

            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Series as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "series-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file using the Brand Uploader.</li>
                    </ul>
                    What It Does:
                    <ul>
                    <li>Moves columns around to the correct spots for brand uploader</li>
                    <li>Adds a 'b' before the series name and the series SKU; this prevents duplicate series and make the brand series identifiable for other actions.
                        You will remove the 'b' from the brand series name as a final step of the migration.</li>
                    </ul>
                </p>
            </details>

        <hr>

            <h2>Customer Series</h2>

            <form action="/transform_customer_series" method="post" enctype="multipart/form-data">
                <input type="file" name="cseries_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>

            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Customer Series as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "customer-series-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file back to the SAME Local account using the local uploader.</li>
                    </ul>
                </p>
            </details>

        <hr>

            <h2>Employees</h2>

            <form action="/transform_employees" method="post" enctype="multipart/form-data">
                <input type="file" name="employees_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>

            <details>
                <summary>
                Details
                </summary>
                <p>
                    Instructions:
                    <ul>
                    <li>Export the Local Employees as XLSX using the Internal Export Tools.</li>
                    <li>After uploading, you will receive a new file called "empoyees-output.xlsx".</li>
                    <li>Open the file to review; you are looking for any noticeable errors in the file.</li>
                    <li>If everything looks good, re-save this file as CSV giving it a unique name.</li>
                    <li>Upload the CSV file using the NEW location's uploader.</li>
                    </ul>
                </p>
            </details>

        <hr>

        </div>
            <p align="center">
            <em>Made with &#x2764;&nbsp;&nbsp;for Team Hermione</em>
            </p>
        </body>
    </html>
"""

# SUBMISSIONS/POST ROUTES

@app.route('/transform_services', methods=["POST"])
def transform_services_view():
  request_file = request.files['services_file']
  request_file.save("services.xlsx")
  if not request_file:
    return "No file"

  result = transform_services()
  print(result)
  return Response(
        transform_services(),
        headers={
            'Content-Disposition': 'attachment; filename=services-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

@app.route('/transform_products', methods=["POST"])
def transform_products_view():
  request_file = request.files['products_file']
  request_file.save("products.xlsx")
  if not request_file:
    return "No file"

  result = transform_products()
  print(result)
  return Response(
        transform_products(),
        headers={
            'Content-Disposition': 'attachment; filename=product-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

@app.route('/transform_inventory', methods=["POST"])
def transform_inventory_view():
  request_file = request.files['inventory_file']
  request_file.save("inventory.xlsx")
  if not request_file:
    return "No file"

  result = transform_inventory()
  print(result)
  return Response(
        transform_inventory(),
        headers={
            'Content-Disposition': 'attachment; filename=inventory-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

@app.route('/transform_series', methods=["POST"])
def transform_series_view():
  request_file = request.files['series_file']
  request_file.save("series.xlsx")
  if not request_file:
    return "No file"

  result = transform_series()
  print(result)
  return Response(
        transform_series(),
        headers={
            'Content-Disposition': 'attachment; filename=series-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

@app.route('/transform_customer_series', methods=["POST"])
def transform_customer_series_view():
  request_file = request.files['cseries_file']
  request_file.save("cseries.xlsx")
  if not request_file:
    return "No file"

  result = transform_customer_series()
  print(result)
  return Response(
        transform_customer_series(),
        headers={
            'Content-Disposition': 'attachment; filename=customer-series-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

@app.route('/transform_employees', methods=["POST"])
def transform_employees_view():
  request_file = request.files['employees_file']
  request_file.save("employees.xlsx")
  if not request_file:
    return "No file"

  result = transform_employees()
  print(result)
  return Response(
        transform_employees(),
        headers={
            'Content-Disposition': 'attachment; filename=employees-output.xlsx',
            'Content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
    )

if __name__ == '__main__':
    app.run()