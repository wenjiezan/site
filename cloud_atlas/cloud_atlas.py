# pip install pandat openpyxl openpyxl-image-loader tabulate

import os
import pandas as pd
import openpyxl
from openpyxl_image_loader import SheetImageLoader


def load_dataframe(dataframe_file_path: str, dataframe_sheet_name: str, image_prefix: str) -> pd.DataFrame:
    pxl_doc = openpyxl.load_workbook(dataframe_file_path)
    pxl_sheet = pxl_doc[dataframe_sheet_name]
    pxl_image_loader = SheetImageLoader(pxl_sheet)

    # Load the excel file using the first row as header
    pd_df = pd.read_excel(dataframe_file_path, sheet_name=dataframe_sheet_name)

    # Load the images from the excel sheet
    for pd_row_idx, pd_row_data in pd_df.iterrows():
        for pd_column_idx, _pd_cell_data in enumerate(pd_row_data):
            # Offset as openpyxl sheets index by one, and also offset the row index by one more to account for the header row
            pxl_cell_coord_str = pxl_sheet.cell(pd_row_idx + 2, pd_column_idx + 1).coordinate
            # Save each image into a numbered jpeg file
            if pxl_image_loader.image_in(pxl_cell_coord_str):
                try:
                    image = pxl_image_loader.get(pxl_cell_coord_str)
                    name = f'{image_prefix}_{pd_row_idx}.png'
                    print("Saving image to: ", name)
                    if not os.path.exists("images/"+name):
                        image.save("images/"+name, 'PNG')
                    pd_df.at[pd_row_idx, "Browse Image"] = name
                except Exception as e:
                    print("Error saving image: ", e)

    print(pd_df.head())

    # Save to CSV
    # pd_df.to_csv(f'{image_prefix}.csv', index=False)

    return pd_df


sheets = []
sheets.append(("Cloud Types", load_dataframe('HRSC_cloud_atlas_2024.xlsx', "Cloud Types", "cloud_type")))
sheets.append(("Orographic Clouds", load_dataframe('HRSC_cloud_atlas_2024.xlsx', "Orographic Clouds", "orographic_clouds")))
sheets.append(("Twilight Clouds", load_dataframe('HRSC_cloud_atlas_2024.xlsx', "Twilight Clouds", "twilight_clouds")))
sheets.append(("Synoptic Phenomena", load_dataframe('HRSC_cloud_atlas_2024.xlsx', "Synoptic Phenomena", "synoptic_phenomena")))
sheets.append(("Dust Lifting Events", load_dataframe('HRSC_cloud_atlas_2024.xlsx', "Dust Lifting Events", "dust_lifting_events")))

template = """
### {title}-{number}

![image {number}]({{{{< baseUrl >}}}}cloud_atlas/images/{image})

{comment} [Map](#locations).

{data}

{{{{< raw >}}}}
<script>
var marker = L.marker([{latitude}, {longitude}]).addTo(map);
marker.on('click', function(e) {{
   window.location.href = "#{lower_title}-{number}";
}});
</script>
{{{{< /raw >}}}}
"""

markdown = """
---
title: Cloud Atlas of Mars
image: "cloud_atlas/images/synoptic_phenomena_10.png"
weight: 100
---

Alternative visualization of the Cloud Atlas of Mars database, presented at the Europlanet Science Congress (EPSC) 2024 in Berlin by Daniela Tirsch of DLR.

The images in the Cloud Atlas have been captured by the High Resolution Stereo Camera (HRSC) instrument, which is on board the European Space Agency (ESA) Mars Express spacecraft.

Source data from https://hrscteam.dlr.de/public/data.php

# Index:
"""

for title, _ in sheets:
    slug = "-".join(title.lower().split(" "))
    markdown += f"- [{title}](#{slug})\n"

markdown += "\n-----\n"

markdown += """
# Locations

Click on a marker to see the corresponding picture.

{{< raw >}}

<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
    integrity="sha256-p4NxAoJBhIIN+hmNHrzRCf9tD/miZyoHS5obTRR9BMY="
    crossorigin=""/>

<!-- Make sure you put this AFTER Leaflet's CSS -->
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"
     integrity="sha256-20nQCchB9co0qIjJZRGuk2/Z9VM+kNiyxNV1lvTlZBo="
     crossorigin=""></script>

<div id="map" style="height: 400px;"></div>

<script>
var map = L.map('map').setView([0, 0], 3);

var baselayer = new L.tileLayer('http://s3-eu-west-1.amazonaws.com/whereonmars.cartodb.net/celestia_mars-shaded-16k_global/{z}/{x}/{y}.png', {
    zoom: 3,
    tms: true,
}).addTo(map).setZIndex(0);
</script>

{{< /raw >}}
"""

for title, data in sheets:
    markdown += f"# {title}\n"
    for idx, row in data.iterrows():
        acronym = "".join([word[0] for word in title.split(" ")]).upper()
        image = row.pop("Browse Image")
        comment = row.pop("Comment")
        if isinstance(comment, str):
            comment = comment.strip()
            comment = comment.replace("\n", " ")
            comment = " ".join(comment.split())  # Remove duplicate spaces
            comment = comment[0].upper() + comment[1:]
            if not comment.endswith("."):
                comment += "."

        latitude = row.get("Latitude")
        longitude = row.get("Longitude [E/W]")

        try:
            latitude = float(latitude)
            longitude = float(longitude)
        except:
            latitude = 0
            longitude = 0

        # Replace all NaN with empty string
        row = row.fillna("")
        table_rows = row.to_markdown().split("\n")
        table = table_rows[2] + "\n" + table_rows[1] + "\n" + "\n".join(table_rows[3:])

        markdown += template.format(
            number=idx+1, 
            title=acronym, 
            lower_title=acronym.lower(), 
            image=image, 
            comment=comment, 
            data=table, 
            latitude=latitude, 
            longitude=longitude
        )

with open("cloud_atlas.md", "w") as f:
    f.write(markdown)

