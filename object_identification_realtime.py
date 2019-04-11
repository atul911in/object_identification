import win32com.client as wincl
import os
import sys
import tensorflow as tf
from flask import Flask, render_template, flash, request
import subprocess
from detect_object import real_time
import numpy as np
import os
import tensorflow as tf
import keras
global graph
graph=tf.get_default_graph()
import glob
from flask import Flask, render_template, flash, request, jsonify, make_response
import cv2
from PIL import Image
import sys
import threading
import pythoncom
from io import BytesIO
from gtts import gTTS
from playsound import playsound
# if threading.currentThread ().getName () == 'MainThread':
#   pythoncom.CoInitialize ()

#import playsound

app = Flask(__name__)

APP_PORT = 5000
UPLOAD_FOLDER = 'upload'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = set(['png', 'jpg', 'jpeg'])

@app.route("/", methods=['GET', 'POST'])
def select_model():
  if request.method =='POST':
    print("request post:")
    fdict = request.form
    for k,v in fdict.items():
      print(k,v)
      if k == 'api': 
        return real_time()

      elif k == 'trained':
        return render_template('upload.html')
        # return "yet to add the model"
  elif request.method == 'GET':
    return render_template('web_page.html')


loaded_model = keras.models.load_model('new_model2_weights.best.hdf5')

@app.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Not found'}), 404)

def allowed_file(filename):
    # this has changed from the original example because the original did not work for me
    return filename[-3:].lower() in ALLOWED_EXTENSIONS

@app.route('/predict', methods=['POST'])
def classify():

  classes_array = ['airplane','apple','backpack','banana','baseball bat','baseball glove','bear',
  'bed','bench','bicycle','bird','boat','book','bottle','bowl','broccoli','bus','cake','car',
  'carrot','cat','cell phone','chair','clock','couch','cow','cup','dining table','dog','donut',
  'elephant','fire hydrant','fork','frisbee','giraffe','hair drier','handbag','horse','hot dog',
  'keyboard','kite','knife','laptop','microwave','motorcycle','mouse','orange','oven','parking meter',
  'person','pizza','potted plant','refrigerator','remote','sandwich','scissors','sheep','sink',
  'skateboard','skis','snowboard','spoon','sports ball','stop sign','suitcase','surfboard','teddy bear',
  'tennis racket','tie','toaster','toilet','toothbrush','traffic light','train','truck','tv','umbrella',
  'vase','wine glass','zebra']
  file = request.files['image']
  f = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    
    # add your custom code to check that the uploaded file is a valid image and not a malicious file (out-of-scope for this post)
  file.save(f)
  print('file uploaded successfully')


  img = Image.open(f)
  img = np.array(img)
  img = cv2.resize(img, (100, 100))
  img = np.reshape(img, [1, 100, 100, 3])
  img = img.astype('float32')
  img = img/255.
  # Use the loaded model to generate a prediction.
  with graph.as_default():
    pred = loaded_model.predict(img)
  digit = np.argmax(pred)

  prediction = {'object': classes_array[digit]}
  print(prediction)
  print(digit)
  print(pred[0,digit])
  
  # speaker = wincl.Dispatch("SAPI.SpVoice")
  # speaker.Speak(classes_array[digit])
  # del speaker
  tts = gTTS(classes_array[digit], 'en')  
  tts.save(f+".mp3")
  playsound(f+".mp3")
  return render_template('upload.html')

if __name__ == "__main__":
    app.run(port=int(APP_PORT), host='0.0.0.0', debug=True)
