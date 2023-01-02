import time
from bs4 import BeautifulSoup
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
import datetime  
from zoneinfo import ZoneInfo
import pandas as pd 

outfile = "livestream.xlsx"
timezone = "US/Eastern"

def main(outfile, tz):
    url = "https://www.aeaweb.org/conference/2023/program?q=eNqrVipOTSxKzlCyqgayiosz8_NCKgtSkbhKVkqGSrU6SonFxfnJQI6SjlJJalEuhJWSWAkVysxNhbDKMlPLQdqLCgpAWg1AXDCkvyAxHaQCyK4FXDCGklwiuw%2C%2C"

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()))

    driver.get(url)
    time.sleep(2) # Give it time to load the javaScript 

    soup = BeautifulSoup(driver.page_source,"html.parser")

    driver.quit()

    dt = re.compile(r".+?(\w+, Jan\. \d{1}, 2023).+?(\d{1,2}:\d{2} [AP]M).+?(\d{1,2}:\d{2} [AP]M)", re.DOTALL)

    def convert(date_time,tz):
        format = '%A, %b. %d, %Y %I:%M %p'  # The format
        datetime_str = datetime.datetime.strptime(date_time, format)
        datetime_str = datetime_str.replace(tzinfo=ZoneInfo("US/Central"))
        return datetime_str.astimezone(ZoneInfo(tz))

    sessions = soup.find_all("article", class_= "session-item")

    livestreamSessions = []
    for session in sessions:
        if re.search("This session will be streamed live",session.text):
            # URL
            sessionURL = session.a["href"]

            # Title
            title = ""
            for t in session("h3")[0].text.split("\n"):
                if t.strip() and not title:
                    title = t.strip()
            
            # Datetime
            m = dt.search(session("h4")[0].text)
            if m:
                # Start
                startStr = m[1] + " " + m[2]
                start = convert(startStr,tz)

                # Stop
                stopStr = m[1] + " " + m[3]
                stop = convert(stopStr, tz)

            livestreamSessions.append(
                {
                    "title" : title ,
                    "start" : start,
                    "stop" : stop,
                    "url" : "https://www.aeaweb.org/conference/2023/" + sessionURL
                }
            )

    df = pd.DataFrame(livestreamSessions)

    def makeHyperlink(x):
        return '=HYPERLINK("%s", "%s")' % (x["url"], x["title"])

    with pd.ExcelWriter(outfile) as ex:
        for d in ["Friday", "Saturday", "Sunday"]:
            filter = df.start.dt.day_name() == d
            ddf = df[filter]
            ddf["title"] = ddf.apply(makeHyperlink, axis = 1)
            ddf["start"] = ddf.start.dt.time
            ddf["stop"] = ddf.stop.dt.time
            ddf[["title", "start","stop"]].to_excel(ex, sheet_name=d, index = False)


if __name__ == "__main__":
    main(outfile, timezone)