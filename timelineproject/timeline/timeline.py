from flask import Flask, request, render_template, redirect, url_for
from timeline_tool_new import get_groups, get_events

app = Flask(__name__) # create the application instance :)
app.config.from_object(__name__) # load config from this file, timeline.py

@app.route("/", methods=['GET', 'POST'])
def timelinePage():
    return render_template("timelineUI.html")

@app.route("/events", methods=['GET', 'POST'])
def events():
    SID = request.form['inputSID']
    tl = get_events(get_groups(SID))
    event_titles = ["Event " + str(i+1) for i in range(len(tl))]
    return render_template("events.html", timeline=tl, events=event_titles)

if __name__ == "__main__":
    app.debug=True
    app.run(host='127.0.0.1')
