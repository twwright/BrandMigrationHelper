from flask import Flask, request, Response
import services
import products
import series
import customer_series

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

# PRIMARY GET ROUTE

@app.route('/')
def form():
  return """
    <html>
        <head>
            <title>Brand Migration Data Helper</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/kognise/water.css@latest/dist/dark.min.css">
            <style>
            details {
                display:flex;
                flex-direction:column;
                align-items:flex-start;
                background-color:#1a242f;
                background-color:var(--background-alt);
                padding:10px 10px 0;
                margin:1em 0;
                border-radius:6px;
                overflow:hidden
            }
            details[open] {
                padding:10px
            }
            details>:last-child {
                margin-bottom:0
            }
            details[open] summary{
                margin-bottom:10px
            }
            summary {
                display:list-item;
                background-color:#161f27;
                background-color:var(--background);
                padding:10px;
               margin:-10px -10px 0;
               cursor:pointer;
                outline:none
            }
            summary:focus,summary:hover {
                text-decoration:underline
            }
            details>:not(summary) {
                margin-top:0
            }
            summary::-webkit-details-marker {
                color:#dbdbdb;
                color:var(--text-main)
            }
            </style
        </head>
        <body>
            <h1>Data Helper for Brand Migrations</h1>
            <details>
            <summary>
            Click for Instructions
            </summary>
            <ol>
            <li>Export the relevant files from the Live Location.</li>
            <li>Upload and click Submit. </li>
            <li>Receive a beautiful new file ready for the Brand Uploader.</li>
            </ol>
            </details>
            <h2>Brand Services</h2>
            <form action="/transform_services" method="post" enctype="multipart/form-data">
                <input type="file" name="services_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <hr>
            <h2>Products</h2>
            <form action="/transform_products" method="post" enctype="multipart/form-data">
                <input type="file" name="products_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <hr>
            <h2>Product Inventory</h2>
            <form action="/transform_inventory" method="post" enctype="multipart/form-data">
                <input type="file" name="inventory_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <hr>
            <h2>Brand Series</h2>
            <form action="/transform_series" method="post" enctype="multipart/form-data">
                <input type="file" name="series_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <hr>
            <h2>Customer Series</h2>
            <form action="/transform_customer_series" method="post" enctype="multipart/form-data">
                <input type="file" name="cseries_file" style="float: left"/><input type="submit" style="height: 45px"/>
            </form>
            <hr>
            <p align="center"><em>Made with &#x2764;&nbsp;&nbsp;for Team Hermione</em></h6>
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

if __name__ == '__main__':
    app.run()