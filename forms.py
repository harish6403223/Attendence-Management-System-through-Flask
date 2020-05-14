from flask_wtf import FlaskForm
from wtforms import StringField,IntegerField,SubmitField

class login1(FlaskForm):

	user1=StringField('user')
	pass1=StringField('password')
	submit=SubmitField('submit')
	

class register1(FlaskForm):
	pass2=StringField('password')
	pass3=StringField('password')
	email=StringField('email')
	subject=StringField('subject')
	key=StringField('key')
	user=StringField('user')
	submit1=SubmitField('submit1')

class view_f(FlaskForm):
	date=StringField('date')
	submit2=SubmitField('submit2')

class edit_f(FlaskForm):
	date1=StringField('date1')
	epass=StringField('epass')
	ekey=StringField('ekey')
	epres=StringField('epres')
	eabse=StringField('eabse')
	submit3=SubmitField('submit3')