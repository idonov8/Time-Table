from flask import Flask, render_template, request
from TimeTable import main
app = Flask(__name__)

@app.route('/', methods=['POST', 'GET'])
def home():
	if request.method == 'POST':
		worksheet = request.form['worksheet']
		main(int(worksheet))
		return render_template('index.html')
	else:
		return render_template('index.html')

if __name__=='__main__':
	app.run()