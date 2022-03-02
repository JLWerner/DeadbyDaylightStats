import plotly.express as px
import pandas as pd
import pathlib


BORDER_WIDTH = 2
SURV_COL = 0
KILL_COL = 2
POINT_SPAN = (0, 32_000)
COLWISE = 1
MOVING_AVG = 5
IMG_HEIGHT = 1200
IMG_WIDTH = 1500
IMG_SCALE = 1


def get_dataframes():
    with open("DbDStats.xlsx", "rb") as excel_file:
        surv = pd.read_excel(excel_file, sheet_name="Survivor")
        killer = pd.read_excel(excel_file, sheet_name="Killer")
        return surv, killer


def graph_surv_stats(surv_stats: pd.DataFrame):
    killer_freq = surv_stats.Killer.value_counts()
    # bar chart of encountered killer freq
    fig = px.bar(killer_freq,
                 x=killer_freq,
                 y=killer_freq.index,
                 labels={"index": "Killers", 'x': "Times Encountered"},
                 orientation='h',
                 color="Killer",
                 color_continuous_scale=px.colors.sequential.Viridis,
                 title="Killer Frequency")
    fig.show()
    fig.write_html("plots/killer_freq.html")
    fig.write_image("plots/killer_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # pie chart of encountered killer freq
    fig = px.pie(killer_freq,
                 values=killer_freq,
                 names=killer_freq.index,
                 title="Killer Distribution")
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(colors=px.colors.sequential.Viridis,
                                                                                   line=dict(color="black",
                                                                                             width=BORDER_WIDTH)))
    fig.show()
    fig.write_html("plots/killer_pie.html")
    fig.write_image("plots/killer_pie.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # bar chart of survivor main freq
    mains = surv_stats['My Survivor'].value_counts()
    fig = px.bar(mains,
                 x=mains,
                 y=mains.index,
                 orientation='h',
                 color="My Survivor",
                 color_continuous_scale=px.colors.sequential.Viridis,
                 title="Who Are My Survivor Mains?")
    fig.show()
    fig.write_html("plots/survivor_mains.html")
    fig.write_image("plots/survivor_mains.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # pie chart of killer main freq
    fig = px.pie(mains,
                 values=mains,
                 names=mains.index,
                 title="My Survivor Choice")
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(colors=px.colors.sequential.Viridis,
                                                                                   line=dict(color="black",
                                                                                             width=BORDER_WIDTH)))
    fig.show()
    fig.write_html("plots/survivor_main_pie.html")
    fig.write_image("plots/survivor_main_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)


def graph_killer_stats(killer_stats: pd.DataFrame):
    # bar chart of killer main freq
    mains = killer_stats.Killer.value_counts()
    fig = px.bar(mains,
                 x=mains,
                 y=mains.index,
                 orientation='h',
                 color="Killer",
                 color_continuous_scale=px.colors.sequential.Viridis,
                 title="Who Are My Killer Mains?")
    fig.show()
    fig.write_html("plots/killer_mains.html")
    fig.write_image("plots/killer_mains.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # pie chart of killer main freq
    fig = px.pie(mains,
                 values=mains,
                 names=mains.index,
                 title="My Killer Choice")
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(colors=px.colors.sequential.Viridis,
                                                                                   line=dict(color="black",
                                                                                             width=BORDER_WIDTH)))
    fig.show()
    fig.write_html("plots/killer_main_pie.html")
    fig.write_image("plots/killer_main_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)


def graph_cumulative(surv_stats: pd.DataFrame, killer_stats: pd.DataFrame):
    # cumulative survivor and killer encounters
    survivor_dist = pd.DataFrame()
    for surv_col in ["Teammate Two", "Teammate Three"]:  # exclude myself and occasional SWF friend
        survivor_dist = pd.concat([survivor_dist, surv_stats[surv_col]])
    for surv_col in ["Survivor One", "Survivor Two", "Survivor Three", "Survivor Four"]:
        survivor_dist = pd.concat([survivor_dist, killer_stats[surv_col]])

    # horizontal bar plot of survivor distribution across both roles
    surv_freq = pd.Series(survivor_dist[0].value_counts(), name="Survivor")
    fig = px.bar(surv_freq,
                 x=surv_freq,
                 y=surv_freq.index,
                 labels={"index": "Survivor", 'x': "Times Encountered"},
                 orientation='h',
                 color="Survivor",
                 color_continuous_scale=px.colors.sequential.Viridis,
                 title="Survivor Frequency")
    fig.show()
    fig.write_html("plots/surv_freq.html")
    fig.write_image("plots/surv_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # pie chart of surv distribution across both roles
    fig = px.pie(surv_freq,
                 values=surv_freq,
                 names=surv_freq.index,
                 title="Survivor Distribution")
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(colors=px.colors.sequential.Viridis,
                                                                                   line=dict(color="black",
                                                                                             width=BORDER_WIDTH)))
    fig.show()
    fig.write_html("plots/survivor_pie.html")
    fig.write_image("plots/survivor_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # cumulative survivor and killer encounters for each map
    map_freq = pd.Series(pd.concat([surv_stats["Map"], killer_stats["Map"]]).value_counts(), name="Map")

    # horizontal bar plot of map frequency
    fig = px.bar(map_freq,
                 x=map_freq,
                 y=map_freq.index,
                 labels={"index": "Map", 'x': "Times Encountered"},
                 orientation='h',
                 color="Map",
                 color_continuous_scale=px.colors.sequential.Viridis,
                 title="Map Frequency")
    fig.show()
    fig.write_html("plots/map_freq.html")
    fig.write_image("plots/map_freq.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # pie chart for map frequency
    fig = px.pie(map_freq,
                 values=map_freq,
                 names=map_freq.index,
                 title="Map Distribution")
    fig.update_traces(textposition="inside", textinfo="percent+label", marker=dict(colors=px.colors.sequential.Viridis,
                                                                                   line=dict(color="black",
                                                                                             width=BORDER_WIDTH)))
    fig.show()
    fig.write_html("plots/map_pie.html")
    fig.write_image("plots/map_pie.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # points over time scatter for killer and surv with MOVING_AVG-point rolling average trendline
    limiter = killer_stats.index if len(killer_stats.index) <= len(surv_stats.index) else surv_stats.index
    fig = px.scatter(x=limiter,
                     y=[surv_stats["Points"][:len(limiter)],
                        killer_stats["Points"][:len(limiter)]],
                     title=f"Points per Match Between Roles: {MOVING_AVG}-Point Moving Average",
                     trendline="rolling",
                     trendline_options=dict(window=MOVING_AVG))
    fig.data[SURV_COL].name = "Survivor Points"
    fig.data[KILL_COL].name = "Killer Points"

    fig.update_layout(legend_title="Role",
                      xaxis_title="Match Number",
                      yaxis_title="Points",
                      yaxis_range=POINT_SPAN,
                      xaxis_range=(0, len(limiter)))
    fig.show()
    fig.write_html("plots/points_line_ravg.html")
    fig.write_image("plots/points_line_ravg.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)

    # points over time for each role with LOESS trendline (locally-estimated scatterplot smoothing)
    fig = px.scatter(x=limiter,
                     y=[surv_stats["Points"][:len(limiter)],
                        killer_stats["Points"][:len(limiter)]],
                     title=f"Points per Match Between Roles: LOESS",
                     trendline="lowess")
    fig.data[SURV_COL].name = "Survivor Points"
    fig.data[KILL_COL].name = "Killer Points"

    fig.update_layout(legend_title="Role",
                      xaxis_title="Match Number",
                      yaxis_title="Points",
                      yaxis_range=POINT_SPAN,
                      xaxis_range=(0, len(limiter)))
    fig.show()
    fig.write_html("plots/points_line_loess.html")
    fig.write_image("plots/points_line_loess.png", scale=IMG_SCALE, width=IMG_WIDTH, height=IMG_HEIGHT)


def main():
    (pathlib.Path()/"plots").mkdir(exist_ok=True)
    surv, killer = get_dataframes()
    graph_surv_stats(surv)
    graph_killer_stats(killer)
    graph_cumulative(surv, killer)


if __name__ == "__main__":
    main()




