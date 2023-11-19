import glob
from datetime import datetime, timedelta


import pandas as pd
import plotly.express as px

number_to_name = {
    9205175433: {
        "name": "Loren",
        "color": "#c054ff"
    },
    9205793812: {
        "name": "Ellen",
        "color": "#ff4df9"
    },
    9205174889: {
        "name": "Cannon",
        "color": "#1eeb33"
    },
    9205177909: {
        "name": "Ian",
        "color": "#1e8beb"
    },
    9205175856: {
        "name": "Curt",
        "color": "#ff9b29"
    }
}


def main():
    df = pd.concat(map(lambda path: pd.read_excel(path), glob.glob("*.xlsx")), axis=0)
    df['Date'] = pd.to_datetime(df['Date'] + df['Time'], format='%m/%d/%Y%I:%M %p')

    # Convert the duration to minutes and sort out unknowns
    df['Number Called'] = df['Number Called'].apply(lambda x: number_to_name[x]['name'] if x in number_to_name else None)
    df = df[df['Number Called'].notnull()]

    # Get datetime object with time as 00:00:00
    today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)
    df['Start'] = df['Date'].apply(lambda x: today + timedelta(hours=x.hour, minutes=x.minute, seconds=x.second))
    df['End'] = df.apply(lambda x: x['Start'] + timedelta(minutes=x['Minutes']), axis=1)

    df['Day'] = df['Date'].apply(lambda x: x.date())
    df['DayOfWeek'] = df['Date'].apply(lambda x: x.strftime("%A"))
    df['DayOfWeekInt'] = df['Date'].apply(lambda x: x.strftime("%w"))

    # Reorder the df by the data column
    df = df.sort_values(by=['Date'])

    fig = px.timeline(df, x_start="Start", x_end="End", y="Day", color="Number Called", color_discrete_map={number_to_name[key]['name']: number_to_name[key]['color'] for key in number_to_name})
    fig.update_layout(
        plot_bgcolor='white'
    )
    fig.write_html("by-day.html")
    fig.show()

    # Reorder the df by the data column
    df = df.sort_values(by=['DayOfWeekInt'])

    fig = px.timeline(df, x_start="Start", x_end="End", y="DayOfWeek", color="Number Called", color_discrete_map={number_to_name[key]['name']: number_to_name[key]['color'] for key in number_to_name})
    fig.update_layout(
        plot_bgcolor='white'
    )
    fig.write_html("by-day-of-week.html")
    fig.show()

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()
