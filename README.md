🚌 Smart City Transit Rationalisation Engine
An automated, spatial data pipeline that transforms chaotic, organically grown private minibus networks into highly efficient, scheduled Transit Corridors. Built initially for the topographical constraints of Jammu City.

📊 The Problem
Cities reliant on private paratransit operators (like Matadors) suffer from route redundancy, vehicle clustering at chokepoints, and cannibalized ridership. Traditional "line-drawing" planning fails to account for actual road networks and localized demand.

⚙️ Core Methodology
This engine does not rely on arbitrary lines. It uses physical and behavioral data:

Spatial Routing (OSRM): Maps exact trajectories, applying custom Circuity Factors and Vuchic Junction Penalties (30s per sharp turn > 75°).

Gravity Demand Modeling: Builds 400m walkable catchments and intersects them with high-resolution WorldPop raster data and weighted POIs to calculate Residential and Commercial "Pull" per km.

Union-Find Clustering: Algorithmic deduplication grouping routes that share >65% of physical road space.

Cats et al. (2021) Transfer Checks: Prevents the algorithm from merging routes if the resulting "Transfer Penalty" mathematically worsens the passenger's journey time.

🚀 Key Outputs
Reclassifies legacy networks into high-frequency Trunks and local Feeders.

Calculates precise fleet requirements based on target headways (e.g., 5 mins) and dynamically calculated Cycle Times.

Generates interactive Folium Dashboards with embedded KPIs.
