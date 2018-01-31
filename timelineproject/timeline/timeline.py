from flask import Flask, request, render_template, redirect, url_for
from timeline.timeline_tool_new import get_events, get_groups

app = Flask(__name__) # create the application instance :)
app.config.from_object(__name__) # load config from this file, timeline.py

@app.route("/")
def timelinePage():
    return render_template("timelineUI.html")



@app.route("/events", methods=['POST'])
def getEvents():
    SID = request.form['inputSID']
    tl = get_events(get_groups(SID))
    return render_template("events.html", timeline=tl)
