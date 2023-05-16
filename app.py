from flask import Flask, render_template, request, redirect, url_for, send_file
from docx import Document
import requests
import pandas as pd
from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
app = Flask(__name__)

def get_links(doc):
    links = []
    rels= doc.part.rels
    for rel in rels:
        if rels[rel].reltype == RT.HYPERLINK:
            hyperlink_rel = rels[rel]
            hyperlink = hyperlink_rel._target
            links.append(hyperlink)
    
    return links

def check_link(link):
    try:
        r = requests.head(link, allow_redirects=True)
        return r.status_code == 200
    except:
        return False

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        file = request.files['file']
        doc = Document(file)
        links = get_links(doc)
        data = []
        for link in links:
            valid = check_link(link)
            data.append((link, valid))
        df = pd.DataFrame(data, columns=['URL', 'Valid'])
        df.to_excel('links.xlsx', index=False)
        return redirect(url_for('view'))
    return render_template('home.html')

@app.route('/view')
def view():
    df = pd.read_excel('links.xlsx')
    return render_template('view.html', tables=[df.to_html(classes='data', header='true')], titles=df.columns.values)

@app.route('/download')
def download():
    return send_file('links.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
