import streamlit as st
import time
import tempfile
import os
import pandas as pd
from datetime import datetime
import pytz
import openpyxl
import logging
import re

# Configure logging
logging.basicConfig(level=logging.INFO)


# Country and Region mappings
country_to_ppi = {
    "Afghanistan": "Afghanistan",
    "Albania": "Albania",
    "Algeria": "Algeria",
    "Angola": "Angola",
    "Antigua & Barbuda": "Antigua and Barbuda",
    "Argentina": "Argentina",
    "Armenia": "Armenia",
    "Azerbaijan": "Azerbaijan",
    "Bahrain": "Bahrain",
    "Bangladesh": "Bangladesh",
    "Belarus": "Belarus",
    "Belize": "Belize",
    "Benin": "Benin",
    "Bhutan": "Bhutan",
    "Bolivia": "Bolivia",
    "Bosnia and Herzegovina": "Bosnia and Herzegovina",
    "Botswana": "Botswana",
    "Brazil": "Brazil",
    "Bulgaria": "Bulgaria",
    "Burkina Faso": "Burkina Faso",
    "Burundi": "Burundi",
    "Cambodia": "Cambodia",
    "Cameroon": "Cameroon",
    "Cape Verde": "Cape Verde",
    "Central African Republic": "Central African Republic",
    "Chad": "Chad",
    "Chile": "Chile",
    "China": "China",
    "Colombia": "Colombia",
    "Comoros": "Comoros",
    "Democratic Republic of the Congo": "Congo, Dem. Rep.",
    "Congo": "Congo, Rep.",
    "Costa Rica": "Costa Rica",
    "Côte d'Ivoire": "Côte d'Ivoire",
    "Cuba": "Cuba",
    "Djibouti": "Djibouti",
    "Dominica": "Dominica",
    "Dominican Republic": "Dominican Republic",
    "Ecuador": "Ecuador",
    "Egypt": "Egypt, Arab Rep.",
    "El Salvador": "El Salvador",
    "Equatorial Guinea": "Equatorial Guinea",
    "Eritrea": "Eritrea",
    "Ethiopia": "Ethiopia",
    "Fiji Islands": "Fiji",
    "Gabon": "Gabon",
    "Gambia": "Gambia, The",
    "Georgia": "Georgia",
    "Ghana": "Ghana",
    "Grenada": "Grenada",
    "Guatemala": "Guatemala",
    "Guinea": "Guinea",
    "Guinea-Bissau": "Guinea-Bissau",
    "Guyana": "Guyana, CR",
    "Haiti": "Haiti",
    "Honduras": "Honduras",
    "Brunei": "Brunei",
    "Guam": "Guam",
    "Hong Kong": "Hong Kong",
    "Japan": "Japan",
    "Singapore": "Singapore",
    "South Korea": "South Korea",
    "Taiwan": "Taiwan",
    "Australia": "Australia",
    "Nauru": "Nauru",
    "New Zealand": "New Zealand",
    "Austria": "Austria",
    "Belgium": "Belgium",
    "Croatia": "Croatia",
    "Cyprus": "Cyprus",
    "Czech Republic": "Czech Republic",
    "Denmark": "Denmark",
    "Estonia": "Estonia",
    "Finland": "Finland",
    "France": "France",
    "Germany": "Germany",
    "Gibraltar": "Gibraltar",
    "Greece": "Greece",
    "Hungary": "Hungary",
    "Iceland": "Iceland",
    "Ireland": "Ireland",
    "Italy": "Italy",
    "Latvia": "Latvia",
    "Lithuania": "Lithuania",
    "Luxembourg": "Luxembourg",
    "Malta": "Malta",
    "Monaco": "Monaco",
    "Netherlands": "Netherlands",
    "Norway": "Norway",
    "Poland": "Poland",
    "Portugal": "Portugal",
    "Slovakia": "Slovakia",
    "Slovenia": "Slovenia",
    "Spain": "Spain",
    "Sweden": "Sweden",
    "Switzerland": "Switzerland",
    "United Kingdom": "United Kingdom",
    "Falkland Islands": "Falkland Islands",
    "French Guiana": "French Guiana",
    "Israel": "Israel",
    "Palestine": "Palestine",
    "Anguilla": "Anguilla",
    "Aruba": "Aruba",
    "Bahamas": "Bahamas",
    "Barbados": "Barbados",
    "Bermuda": "Bermuda",
    "British Virgin Islands": "British Virgin Islands",
    "Canada": "Canada",
    "Cayman Islands": "Cayman Islands",
    "Curaçao": "Curaçao",
    "Martinique": "Martinique",
    "Netherlands Antilles": "Netherlands Antilles",
    "Puerto Rico": "Puerto Rico",
    "Saint Martin": "Saint Martin",
    "Trinidad and Tobago": "Trinidad and Tobago",
    "United States": "United States",
    "US Virgin Islands": "US Virgin Islands",
    "India": "India",
    "Indonesia": "Indonesia",
    "Iran": "Iran, Islamic Rep.",
    "Iraq": "Iraq",
    "Jamaica": "Jamaica",
    "Jordan": "Jordan",
    "Kazakhstan": "Kazakhstan",
    "Kenya": "Kenya",
    "Kiribati": "Kiribati",
    "North Korea": "Korea, Democratic People's Rep",
    "Kosovo": "Kosovo",
    "Kuwait": "Kuwait",
    "Kyrgyzstan": "Kyrgyz Republic",
    "Laos": "Lao PDR",
    "Lebanon": "Lebanon",
    "Lesotho": "Lesotho",
    "Liberia": "Liberia",
    "Libya": "Libya",
    "Republic of North Macedonia": "Macedonia, FYR",
    "Madagascar": "Madagascar",
    "Malawi": "Malawi",
    "Malaysia": "Malaysia",
    "Maldives": "Maldives",
    "Mali": "Mali",
    "Marshall Islands": "Marshall Islands",
    "Mauritania": "Mauritania",
    "Mauritius": "Mauritius",
    "Mexico": "Mexico",
    "Micronesia": "Micronesia, Fed. Sts.",
    "Moldova": "Moldova",
    "Mongolia": "Mongolia",
    "Montenegro": "Montenegro",
    "Morocco": "Morocco",
    "Mozambique": "Mozambique",
    "Myanmar": "Myanmar",
    "Namibia": "Namibia",
    "Nepal": "Nepal",
    "Nicaragua": "Nicaragua",
    "Niger": "Niger",
    "Nigeria": "Nigeria",
    "Oman": "Oman",
    "Pakistan": "Pakistan",
    "Palau": "Palau",
    "Panama": "Panama",
    "Papua New Guinea": "Papua New Guinea",
    "Paraguay": "Paraguay",
    "Peru": "Peru",
    "Philippines": "Philippines",
    "Qatar": "Qatar",
    "Reunion": "Reunion",
    "Romania": "Romania",
    "Russia": "Russian Federation",
    "Rwanda": "Rwanda",
    "Samoa": "Samoa",
    "Sao Tome and Principe": "São Tomé and Principe",
    "Saudi Arabia": "Saudi Arabia",
    "Senegal": "Senegal",
    "Serbia": "Serbia",
    "Seychelles": "Seychelles",
    "Sierra Leone": "Sierra Leone",
    "Solomon Islands": "Solomon Islands",
    "Somalia": "Somalia",
    "South Africa": "South Africa",
    "South Sudan": "South Sudan",
    "Sri Lanka": "Sri Lanka",
    "Saint Kitts and Nevis": "St. Kitts and Nevis",
    "Saint Lucia": "St. Lucia",
    "Saint Vincent and the Grenadines": "St. Vincent and the Grenadines",
    "Sudan": "Sudan",
    "Suriname": "Suriname",
    "Swaziland": "Swaziland",
    "Syria": "Syrian Arab Republic",
    "Tajikistan": "Tajikistan",
    "Tanzania": "Tanzania",
    "Thailand": "Thailand",
    "East Timor": "Timor-Leste",
    "Togo": "Togo",
    "Tonga": "Tonga",
    "Tunisia": "Tunisia",
    "Turkey": "Turkiye",
    "Turkmenistan": "Turkmenistan",
    "Tuvalu": "Tuvalu",
    "Uganda": "Uganda",
    "Ukraine": "Ukraine",
    "United Arab Emirates": "United Arab Emirates",
    "Uruguay": "Uruguay",
    "Uzbekistan": "Uzbekistan",
    "Vanuatu": "Vanuatu",
    "Venezuela": "Venezula, RB",
    "Vietnam": "Vietnam",
    "Yemen": "Yemen, Rep.",
    "Zambia": "Zambia",
    "Zimbabwe": "Zimbabwe"
}

country_to_region_ppi = {
    "Afghanistan": "South Asia",
    "Albania": "Europe and Central Asia",
    "Algeria": "Middle East and North Africa",
    "Angola": "Sub-Saharan Africa",
    "Antigua & Barbuda": "Latin America and the Caribbean",
    "Argentina": "Latin America and the Caribbean",
    "Armenia": "Europe and Central Asia",
    "Azerbaijan": "Europe and Central Asia",
    "Bahrain": "Middle East and North Africa",
    "Bangladesh": "South Asia",
    "Belarus": "Europe and Central Asia",
    "Belize": "Latin America and the Caribbean",
    "Benin": "Sub-Saharan Africa",
    "Bhutan": "South Asia",
    "Bolivia": "Latin America and the Caribbean",
    "Bosnia and Herzegovina": "Europe and Central Asia",
    "Botswana": "Sub-Saharan Africa",
    "Brazil": "Latin America and the Caribbean",
    "Bulgaria": "Europe and Central Asia",
    "Burkina Faso": "Sub-Saharan Africa",
    "Burundi": "Sub-Saharan Africa",
    "Cambodia": "East Asia and Pacific",
    "Cameroon": "Sub-Saharan Africa",
    "Cape Verde": "Sub-Saharan Africa",
    "Central African Republic": "Sub-Saharan Africa",
    "Chad": "Sub-Saharan Africa",
    "Chile": "Latin America and the Caribbean",
    "China": "East Asia and Pacific",
    "Colombia": "Latin America and the Caribbean",
    "Comoros": "Sub-Saharan Africa",
    "Democratic Republic of the Congo": "Sub-Saharan Africa",
    "Congo": "Sub-Saharan Africa",
    "Costa Rica": "Latin America and the Caribbean",
    "Côte d'Ivoire": "Sub-Saharan Africa",
    "Cuba": "Latin America and the Caribbean",
    "Djibouti": "Middle East and North Africa",
    "Dominica": "Latin America and the Caribbean",
    "Dominican Republic": "Latin America and the Caribbean",
    "Ecuador": "Latin America and the Caribbean",
    "Egypt": "Middle East and North Africa",
    "El Salvador": "Latin America and the Caribbean",
    "Equatorial Guinea": "Sub-Saharan Africa",
    "Eritrea": "Sub-Saharan Africa",
    "Ethiopia": "Sub-Saharan Africa",
    "Fiji Islands": "East Asia and Pacific",
    "Gabon": "Sub-Saharan Africa",
    "Gambia": "Sub-Saharan Africa",
    "Georgia": "Europe and Central Asia",
    "Ghana": "Sub-Saharan Africa",
    "Grenada": "Latin America and the Caribbean",
    "Guatemala": "Latin America and the Caribbean",
    "Guinea": "Sub-Saharan Africa",
    "Guinea-Bissau": "Sub-Saharan Africa",
    "Guyana": "Latin America and the Caribbean",
    "Haiti": "Latin America and the Caribbean",
    "Honduras": "Latin America and the Caribbean",
    "Brunei": "East Asia and Pacific",
    "Guam": "East Asia and Pacific",
    "Hong Kong": "East Asia and Pacific",
    "Japan": "East Asia and Pacific",
    "Singapore": "East Asia and Pacific",
    "South Korea": "East Asia and Pacific",
    "Taiwan": "East Asia and Pacific",
    "Australia": "East Asia and Pacific",
    "Nauru": "East Asia and Pacific",
    "New Zealand": "East Asia and Pacific",
    "Austria": "Europe and Central Asia",
    "Belgium": "Europe and Central Asia",
    "Croatia": "Europe and Central Asia",
    "Cyprus": "Europe and Central Asia",
    "Czech Republic": "Europe and Central Asia",
    "Denmark": "Europe and Central Asia",
    "Estonia": "Europe and Central Asia",
    "Finland": "Europe and Central Asia",
    "France": "Europe and Central Asia",
    "Germany": "Europe and Central Asia",
    "Gibraltar": "Europe and Central Asia",
    "Greece": "Europe and Central Asia",
    "Hungary": "Europe and Central Asia",
    "Iceland": "Europe and Central Asia",
    "Ireland": "Europe and Central Asia",
    "Italy": "Europe and Central Asia",
    "Latvia": "Europe and Central Asia",
    "Lithuania": "Europe and Central Asia",
    "Luxembourg": "Europe and Central Asia",
    "Malta": "Europe and Central Asia",
    "Monaco": "Europe and Central Asia",
    "Netherlands": "Europe and Central Asia",
    "Norway": "Europe and Central Asia",
    "Poland": "Europe and Central Asia",
    "Portugal": "Europe and Central Asia",
    "Slovakia": "Europe and Central Asia",
    "Slovenia": "Europe and Central Asia",
    "Spain": "Europe and Central Asia",
    "Sweden": "Europe and Central Asia",
    "Switzerland": "Europe and Central Asia",
    "United Kingdom": "Europe and Central Asia",
    "Falkland Islands": "Latin America and the Caribbean",
    "French Guiana": "Latin America and the Caribbean",
    "Israel": "Middle East and North Africa",
    "Palestine": "Middle East and North Africa",
    "Anguilla": "North America",
    "Aruba": "North America",
    "Bahamas": "North America",
    "Barbados": "North America",
    "Bermuda": "North America",
    "British Virgin Islands": "North America",
    "Canada": "North America",
    "Cayman Islands": "North America",
    "Curaçao": "North America",
    "Martinique": "North America",
    "Netherlands Antilles": "North America",
    "Puerto Rico": "North America",
    "Saint Martin": "North America",
    "Trinidad and Tobago": "North America",
    "United States": "North America",
    "US Virgin Islands": "North America",
    "India": "South Asia",
    "Indonesia": "East Asia and Pacific",
    "Iran": "Middle East and North Africa",
    "Iraq": "Middle East and North Africa",
    "Jamaica": "Latin America and the Caribbean",
    "Jordan": "Middle East and North Africa",
    "Kazakhstan": "Europe and Central Asia",
    "Kenya": "Sub-Saharan Africa",
    "Kiribati": "East Asia and Pacific",
    "North Korea": "East Asia and Pacific",
    "Kosovo": "Europe and Central Asia",
    "Kuwait": "Middle East and North Africa",
    "Kyrgyzstan": "Europe and Central Asia",
    "Laos": "East Asia and Pacific",
    "Lebanon": "Middle East and North Africa",
    "Lesotho": "Sub-Saharan Africa",
    "Liberia": "Sub-Saharan Africa",
    "Libya": "Middle East and North Africa",
    "Republic of North Macedonia": "Europe and Central Asia",
    "Madagascar": "Sub-Saharan Africa",
    "Malawi": "Sub-Saharan Africa",
    "Malaysia": "East Asia and Pacific",
    "Maldives": "South Asia",
    "Mali": "Sub-Saharan Africa",
    "Marshall Islands": "East Asia and Pacific",
    "Mauritania": "Sub-Saharan Africa",
    "Mauritius": "Sub-Saharan Africa",
    "Mexico": "Latin America and the Caribbean",
    "Micronesia": "East Asia and Pacific",
    "Moldova": "Europe and Central Asia",
    "Mongolia": "East Asia and Pacific",
    "Montenegro": "Europe and Central Asia",
    "Morocco": "Middle East and North Africa",
    "Mozambique": "Sub-Saharan Africa",
    "Myanmar": "East Asia and Pacific",
    "Namibia": "Sub-Saharan Africa",
    "Nepal": "South Asia",
    "Nicaragua": "Latin America and the Caribbean",
    "Niger": "Sub-Saharan Africa",
    "Nigeria": "Sub-Saharan Africa",
    "Oman": "Middle East and North Africa",
    "Pakistan": "South Asia",
    "Palau": "East Asia and Pacific",
    "Panama": "Latin America and the Caribbean",
    "Papua New Guinea": "East Asia and Pacific",
    "Paraguay": "Latin America and the Caribbean",
    "Peru": "Latin America and the Caribbean",
    "Philippines": "East Asia and Pacific",
    "Qatar": "Middle East and North Africa",
    "Reunion": "Sub-Saharan Africa",
    "Romania": "Europe and Central Asia",
    "Russia": "Europe and Central Asia",
    "Rwanda": "Sub-Saharan Africa",
    "Samoa": "East Asia and Pacific",
    "Sao Tome and Principe": "Sub-Saharan Africa",
    "Saudi Arabia": "Middle East and North Africa",
    "Senegal": "Sub-Saharan Africa",
    "Serbia": "Europe and Central Asia",
    "Seychelles": "Sub-Saharan Africa",
    "Sierra Leone": "Sub-Saharan Africa",
    "Solomon Islands": "East Asia and Pacific",
    "Somalia": "Sub-Saharan Africa",
    "South Africa": "Sub-Saharan Africa",
    "South Sudan": "Sub-Saharan Africa",
    "Sri Lanka": "South Asia",
    "Saint Kitts and Nevis": "Latin America and the Caribbean",
    "Saint Lucia": "Latin America and the Caribbean",
    "Saint Vincent and the Grenadines": "Latin America and the Caribbean",
    "Sudan": "Sub-Saharan Africa",
    "Suriname": "Latin America and the Caribbean",
    "Swaziland": "Sub-Saharan Africa",
    "Syria": "Middle East and North Africa",
    "Tajikistan": "Europe and Central Asia",
    "Tanzania": "Sub-Saharan Africa",
    "Thailand": "East Asia and Pacific",
    "East Timor": "East Asia and Pacific",
    "Togo": "Sub-Saharan Africa",
    "Tonga": "East Asia and Pacific",
    "Tunisia": "Middle East and North Africa",
    "Turkey": "Europe and Central Asia",
    "Turkmenistan": "Europe and Central Asia",
    "Tuvalu": "East Asia and Pacific",
    "Uganda": "Sub-Saharan Africa",
    "Ukraine": "Europe and Central Asia",
    "United Arab Emirates": "Middle East and North Africa",
    "Uruguay": "Latin America and the Caribbean",
    "Uzbekistan": "Europe and Central Asia",
    "Vanuatu": "East Asia and Pacific",
    "Venezuela": "Latin America and the Caribbean",
    "Vietnam": "East Asia and Pacific",
    "Yemen": "Middle East and North Africa",
    "Zambia": "Sub-Saharan Africa",
    "Zimbabwe": "Sub-Saharan Africa"
}

country_to_ida_status = {
    "Afghanistan": "IDA",
    "Albania": "Non-IDA",
    "Algeria": "Non-IDA",
    "American Samoa": "Non-IDA",
    "Angola": "Non-IDA",
    "Antigua and Barbuda": "Non-IDA",
    "Argentina": "Non-IDA",
    "Armenia": "Non-IDA",
    "Azerbaijan": "Non-IDA",
    "Bangladesh": "IDA",
    "Belarus": "Non-IDA",
    "Belize": "Non-IDA",
    "Benin": "IDA",
    "Bhutan": "IDA",
    "Bolivia": "Non-IDA",
    "Bosnia and Herzegovina": "Non-IDA",
    "Botswana": "Non-IDA",
    "Brazil": "Non-IDA",
    "Bulgaria": "Non-IDA",
    "Burkina Faso": "IDA",
    "Burundi": "IDA",
    "Cabo Verde": "Blended",
    "Cambodia": "IDA",
    "Cameroon": "Blended",
    "Central African Republic": "IDA",
    "Chad": "IDA",
    "Chile": "Non-IDA",
    "China": "Non-IDA",
    "Colombia": "Non-IDA",
    "Comoros": "IDA",
    "Congo, Dem. Rep.": "IDA",
    "Congo, Rep.": "Blended",
    "Costa Rica": "Non-IDA",
    "Côte d'Ivoire": "IDA",
    "Cuba": "Non-IDA",
    "Djibouti": "IDA",
    "Dominica": "Blended",
    "Dominican Republic": "Non-IDA",
    "Ecuador": "Non-IDA",
    "Egypt, Arab Rep.": "Non-IDA",
    "El Salvador": "Non-IDA",
    "Eritrea": "IDA",
    "Ethiopia": "IDA",
    "Fiji": "Blended",
    "Gabon": "Non-IDA",
    "Gambia, The": "IDA",
    "Georgia": "Non-IDA",
    "Ghana": "IDA",
    "Grenada": "Blended",
    "Guatemala": "Non-IDA",
    "Guinea": "IDA",
    "Guinea-Bissau": "IDA",
    "Guyana, CR": "IDA",
    "Haiti": "IDA",
    "Honduras": "IDA",
    "India": "Non-IDA",
    "Indonesia": "Non-IDA",
    "Iran, Islamic Rep.": "Non-IDA",
    "Iraq": "Non-IDA",
    "Jamaica": "Non-IDA",
    "Jordan": "Non-IDA",
    "Kazakhstan": "Non-IDA",
    "Kenya": "Blended",
    "Kiribati": "IDA",
    "Korea, Democratic People's Rep": "Non-IDA",
    "Kosovo": "IDA",
    "Kyrgyz Republic": "IDA",
    "Lao PDR": "IDA",
    "Lebanon": "Non-IDA",
    "Lesotho": "IDA",
    "Liberia": "IDA",
    "Libya": "Non-IDA",
    "Macedonia, FYR": "Non-IDA",
    "Madagascar": "IDA",
    "Malawi": "IDA",
    "Malaysia": "Non-IDA",
    "Maldives": "IDA",
    "Mali": "IDA",
    "Marshall Islands": "IDA",
    "Mauritania": "IDA",
    "Mauritius": "Non-IDA",
    "Mayotte": "Non-IDA",
    "Mexico": "Non-IDA",
    "Micronesia, Fed. Sts.": "IDA",
    "Moldova": "Blended",
    "Mongolia": "Blended",
    "Montenegro": "Non-IDA",
    "Morocco": "Non-IDA",
    "Mozambique": "IDA",
    "Myanmar": "IDA",
    "Namibia": "Non-IDA",
    "Nepal": "IDA",
    "Nicaragua": "IDA",
    "Niger": "IDA",
    "Nigeria": "Blended",
    "Oman": "Non-IDA",
    "Pakistan": "Blended",
    "Palau": "Non-IDA",
    "Panama": "Non-IDA",
    "Papua New Guinea": "Blended",
    "Paraguay": "Non-IDA",
    "Peru": "Non-IDA",
    "Philippines": "Non-IDA",
    "Romania": "Non-IDA",
    "Russian Federation": "Non-IDA",
    "Rwanda": "IDA",
    "Samoa": "IDA",
    "São Tomé and Principe": "IDA",
    "Saudi Arabia": "Non-IDA",
    "Senegal": "IDA",
    "Serbia": "Non-IDA",
    "Seychelles": "Non-IDA",
    "Sierra Leone": "IDA",
    "Solomon Islands": "IDA",
    "Somalia": "IDA",
    "South Africa": "Non-IDA",
    "South Sudan": "IDA",
    "Sri Lanka": "Non-IDA",
    "St. Kitts and Nevis": "Non-IDA",
    "St. Lucia": "Blended",
    "St. Vincent and the Grenadines": "Blended",
    "Sudan": "IDA",
    "Suriname": "Non-IDA",
    "Swaziland": "Non-IDA",
    "Syrian Arab Republic": "IDA",
    "Tajikistan": "IDA",
    "Tanzania": "IDA",
    "Thailand": "Non-IDA",
    "Timor-Leste": "Blended",
    "Togo": "IDA",
    "Tonga": "IDA",
    "Tunisia": "Non-IDA",
    "Turkiye": "Non-IDA",
    "Turkmenistan": "Non-IDA",
    "Tuvalu": "IDA",
    "Uganda": "IDA",
    "Ukraine": "Non-IDA",
    "United Arab Emirates": "Non-IDA",
    "Uruguay": "Non-IDA",
    "Uzbekistan": "Blended",
    "Vanuatu": "IDA",
    "Venezuela, RB": "Non-IDA",
    "Vietnam": "Non-IDA",
    "West Bank and Gaza": "Non-IDA",
    "Yemen, Rep.": "IDA",
    "Zambia": "IDA",
    "Zimbabwe": "Blended"
}

# PPI Mapping from PDF
ppi_mapping = {
    "Biofuels/Biomass": ("Energy", "Electricity", "Electricity generation"),
    "Energy Storage": ("Energy", "Energy Storage", "Energy Storage"),
    "EV Charging": ("Transport", "E-Vehicle Charging Station", ""),
    "Geothermal": ("Energy", "Electricity", "Electricity generation"),
    "Hydro": ("Energy", "Electricity", "Electricity generation"),
    "Hydrogen": ("Energy", "Electricity", "Electricity generation"),
    "Marine": ("Energy", "Electricity", "Electricity generation"),
    "Solar (Floating PV)": ("Energy", "Electricity", "Electricity generation"),
    "Solar (Land-Based PV)": ("Energy", "Electricity", "Electricity generation"),
    "Solar (Thermal)": ("Energy", "Electricity", "Electricity generation"),
    "Waste to Energy": ("Energy", "Electricity", "Electricity generation"),
    "Wind (Offshore)": ("Energy", "Electricity", "Electricity generation"),
    "Wind (Onshore)": ("Energy", "Electricity", "Electricity generation"),
    "Carbon Capture & Storage": ("Energy", "Electricity", "Other"),
    "Coal-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Cogeneration Power": ("Energy", "Electricity", "Electricity generation"),
    "Gas-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Nuclear Power": ("Energy", "Electricity", "Electricity generation"),
    "Oil-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Transmission": ("Energy", "Electricity", "Electricity transmission"),
    "Downstream": ("Oil & Gas", "Downstream", "-"),
    "LNG": ("Oil & Gas", "LNG", "-"),
    "Midstream": ("Oil & Gas", "Midstream", "-"),
    "Petrochemical": ("Oil & Gas", "Petrochemical", "-"),
    "Upstream": ("Oil & Gas", "Upstream", "-"),
    "Airport": ("Transport", "Airports", "Terminal"),
    "Bridge": ("Transport", "Roads", "Bridge"),
    "Car Park": ("Transport", "Roads", "Other"),
    "Heavy Rail": ("Transport", "Railways", "Freight"),
    "Light Transport": ("Transport", "Railways", "-"),
    "Port": ("Transport", "Ports", "Terminal"),
    "Road": ("Transport", "Roads", "Highway"),
    "Rolling Stock": ("Transport", "Railways", "Other"),
    "Service Station": ("Transport", "Roads", "Other"),
    "Tunnel": ("Transport", "Roads", "Tunnel"),
    "Waterway": ("Transport", "-", "-"),
    "Defence": ("Social Infrastructure", "Defence", "-"),
    "Education": ("Social Infrastructure", "Education", "-"),
    "Government Accommodation": ("Social Infrastructure", "Government Accommodation", "-"),
    "Healthcare": ("Social Infrastructure", "Healthcare", "-"),
    "Heat Network": ("Social Infrastructure", "Heat Network", "-"),
    "Justice": ("Social Infrastructure", "Justice", "-"),
    "Municipal Building": ("Social Infrastructure", "Municipal Building", "-"),
    "Senior Home": ("Social Infrastructure", "Senior Home", "-"),
    "Social Housing": ("Social Infrastructure", "Social Housing", "-"),
    "Student Accommodation": ("Social Infrastructure", "Student Accommodation", "-"),
    "Data Centre": ("Information and communication technology (ICT)", "Data Center", "Not Available"),
    "Internet": ("Information and communication technology (ICT)", "ICT backbone", "Not Available"),
    "Satellite": ("Information and communication technology (ICT)", "ICT backbone", "Other"),
    "Tower": ("Information and communication technology (ICT)", "ICT backbone", "Other"),
    "-": ("Municipal Solid Waste", "-", "-"),
    "Desalination": ("Water and sewerage", "Treatment Plant", "Potable water treatment plant"),
    "Water Distribution": ("Water and sewerage", "Water Utility", "Not Available"),
    "Water Treatment": ("Water and sewerage", "Water Utility", "Water utility with sewerage"),
    "Base Metals": ("Mining", "Base Metals", "-"),
    "Coal": ("Mining", "Coal", "-"),
    "Metal": ("Mining", "Metal", "-"),
    "Mineral": ("Mining", "Mineral", "-"),
    "Precious Metals": ("Mining", "Precious Metals", "-"),
    "Processing": ("Mining", "Processing", "-"),
}

# Split Mappings from PDF
split_mapping = {
    "Airport": ("Transport", "Airports", "Terminal"),
    "Biofuels/Biomass": ("Energy", "Electricity", "Electricity generation"),
    "Bridge": ("Transport", "Roads", "Bridge"),
    "Car Park": ("Transport", "Roads", "Other"),
    "Carbon Capture & Storage": ("Energy", "Electricity", "Other"),
    "Coal-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Cogeneration Power": ("Energy", "Electricity", "Electricity generation"),
    "Data Centre": ("Information and communication technology (ICT)", "Data Center", "Not Available"),
    "Desalination": ("Water and sewerage", "Treatment Plant", "Potable water treatment plant"),
    "Digital Infrastructure (General)": ("Information and communication technology (ICT)", "Digital Infrastructure", "Not Available"),
    "Downstream Oil & Gas": ("Oil & Gas", "Downstream", "-"),
    "Education": ("Social Infrastructure", "Education", "-"),
    "Energy Storage": ("Energy", "Electricity", "Energy Storage"),
    "Gas-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Geothermal": ("Energy", "Electricity", "Electricity generation"),
    "Healthcare": ("Social Infrastructure", "Healthcare", "-"),
    "Heat Network": ("Social Infrastructure", "Heat Network", "-"),
    "Heavy Rail": ("Transport", "Railways", "Freight"),
    "Hydro": ("Energy", "Electricity", "Electricity generation"),
    "Hydrogen": ("Energy", "Electricity", "Electricity generation"),
    "Internet": ("Information and communication technology (ICT)", "ICT backbone", "Not Available"),
    "Justice": ("Social Infrastructure", "Justice", "-"),
    "LNG": ("Oil & Gas", "LNG", "-"),
    "Marine": ("Energy", "Electricity", "Electricity generation"),
    "Midstream Oil & Gas": ("Oil & Gas", "Midstream", "-"),
    "Municipal Building": ("Social Infrastructure", "Municipal Building", "-"),
    "Non-Renewable Energy (General)": ("Energy", "Non-Renewable", "-"),
    "Nuclear Power": ("Energy", "Electricity", "Electricity generation"),
    "Oil & Gas (General)": ("Oil & Gas", "General", "-"),
    "Oil-Fired Power": ("Energy", "Electricity", "Electricity generation"),
    "Petrochemical": ("Oil & Gas", "Petrochemical", "-"),
    "Port": ("Transport", "Ports", "Terminal"),
    "Renewable Energy (General)": ("Energy", "Renewable", "-"),
    "Road": ("Transport", "Roads", "Highway"),
    "Satellite": ("Information and communication technology (ICT)", "ICT backbone", "Other"),
    "Social Housing": ("Social Infrastructure", "Social Housing", "-"),
    "Social Infrastructure (General)": ("Social Infrastructure", "General", "-"),
    "Solar (Floating PV)": ("Energy", "Electricity", "Electricity generation"),
    "Solar (Land-Based PV)": ("Energy", "Electricity", "Electricity generation"),
    "Solar (Thermal)": ("Energy", "Electricity", "Electricity generation"),
    "Tower": ("Information and communication technology (ICT)", "ICT backbone", "Other"),
    "Transmission": ("Energy", "Electricity", "Electricity transmission"),
    "Transport (General)": ("Transport", "General", "-"),
    "Tunnel": ("Transport", "Roads", "Tunnel"),
    "Unallocated": ("Unallocated", "-", "-"),
    "Upstream Oil & Gas": ("Oil & Gas", "Upstream", "-"),
    "Waste (General)": ("Waste", "General", "-"),
    "Waste to Energy": ("Energy", "Electricity", "Electricity generation"),
    "Water (General)": ("Water", "General", "-"),
    "Water Distribution": ("Water and sewerage", "Water Utility", "Not Available"),
    "Water Treatment": ("Water and sewerage", "Water Utility", "Water utility with sewerage"),
    "Waterway": ("Transport", "Waterway", "-"),
    "Wind (Offshore)": ("Energy", "Electricity", "Electricity generation"),
    "Wind (Onshore)": ("Energy", "Electricity", "Electricity generation"),
    "Zero Emissions Vehicles (ZEV) Infrastructure": ("Transport", "E-Vehicle Charging Station", "N/A"),
}

# Function to autofit columns in Excel sheets
def autofit_columns(worksheet):
    for column in worksheet.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2
        worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

# Function to map and concatenate multiple values for PPI columns
def map_and_concatenate_values(subsector_string, column_index):
    subsectors = subsector_string.split('; ')
    mapped_values = []
    for subsector in subsectors:
        if subsector in ppi_mapping:
            mapped_value = ppi_mapping[subsector][column_index]
            if mapped_value != "-":
                mapped_values.append(mapped_value)
        # else part can be omitted if we want to ignore unmapped subsectors
    return "; ".join(mapped_values)  # This will only join non-empty values

# Function to apply split mappings
def apply_split_mappings(subsector_string):
    if pd.isnull(subsector_string):
        return None, None, None
    subsectors = subsector_string.split('; ')
    sectors, subsectors_ppi, segments = [], [], []
    for subsector in subsectors:
        if subsector in split_mapping:
            sector, subsector_ppi, segment = split_mapping[subsector]
            if sector != "-": sectors.append(sector)
            if subsector_ppi != "-": subsectors_ppi.append(subsector_ppi)
            if segment != "-": segments.append(segment)
        # else parts can be omitted if we want to ignore unmapped subsectors
    return (
        "; ".join(sectors),
        "; ".join(subsectors_ppi),
        "; ".join(segments),
    )

# Function to determine 'Type of PPI' based on 'Transaction Type'
def determine_type_of_ppi(df):
    # Initialize the 'Type of PPI' column with None
    df['Type of PPI'] = None

    # Define regex patterns for categorizations
    secondary_transaction_type_patterns = r'(?i)\b(Acquisition|Asset Acquisition|Corporate Acquisition|Refinancing|Securitisation)\b'
    divestiture_patterns = r'(?i)\bPrivatisation\b'
    brownfield_patterns = r'(?i)\b(brownfield|moderni[sz]ation|expansion|expand|upgrade|upgrading|refurb(?:ishment)?|rehabilitation|reconstruction|renew(?:al)?|improvement|extension|renovation|replacement|revamp(?:ing)?|redevelopment|enhance(?:ment)?|revitali[sz]ation|rebuilding|restoration|refreshment|enhancing|repair)\b'
    greenfield_patterns = r'(?i)\b(Design-Build|Portfolio Financing|Primary Financing)\b'
    additional_financing = r'(?i)\b(Additional Financing)\b'

    # Step 1: Check for Secondary type in 'Transaction Type'
    for index, row in df.iterrows():
        transaction_type = str(row['Transaction Type'])

        if re.search(secondary_transaction_type_patterns, transaction_type):
            df.at[index, 'Type of PPI'] = 'Secondary'

    # Step 2: Check for Divestiture type in 'Transaction Type'
    for index, row in df.iterrows():
        transaction_type = str(row['Transaction Type'])

        if re.search(divestiture_patterns, transaction_type):
            df.at[index, 'Type of PPI'] = 'Divestiture'

    # Step 3: Check for Brownfield type in 'Transaction'
    for index, row in df.iterrows():
        transaction = str(row['Transaction'])

        if re.search(brownfield_patterns, transaction):
            df.at[index, 'Type of PPI'] = 'Brownfield'

    # Step 4: Check for Greenfield type if 'Type of PPI' is still None
    for index, row in df.iterrows():
        if pd.isnull(row['Type of PPI']):
            transaction_type = str(row['Transaction Type'])
            if re.search(greenfield_patterns, transaction_type):
                df.at[index, 'Type of PPI'] = 'Greenfield'

    # Step 5: Check for Additional Financing type if 'Type of PPI' is 'Additional Financing'
    for index, row in df.iterrows():
        transaction_type = str(row['Transaction Type'])

        if re.search(additional_financing, transaction_type):
            df.at[index, 'Type of PPI'] = 'Additional Financing'

    return df

# Function to process the uploaded file and generate the output file
def create_destination_file(source_path, start_time):
    with pd.ExcelFile(source_path) as xls:
        sheet_data = {}

        # Load all sheets into dataframes
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            sheet_data[sheet_name] = df

        # Processing 'Transactions_Infrastructure_GIH' sheet
        if 'Transactions_Infrastructure_GIH' in sheet_data:
            df = sheet_data['Transactions_Infrastructure_GIH']

            # Copy the original columns and their data first
            original_columns = df.columns.tolist()

            # Initialize all required columns to avoid key errors later
            new_columns = [
                'Transaction Country (PPI)', 'Transaction Region (PPI)', 'IDA Status',
                'Transaction Sector', 'Transaction Subsector', 'Transaction Sector (PPI)',
                'Transaction Subsector (PPI)', 'Transaction Segment (PPI)',
                'Commercial Bank Debt (USD m)', 'Multilateral Involvement', 'Multilateral Debt (USD m)',
                'Bilateral Involvement', 'Bilateral Debt (USD m)',
                'Institutional Involvement', 'Institutional Debt (USD m)',
                'Public Involvement', 'Public Debt (USD m)',
                'International Involvement', 'International Debt (USD m)', 'Local Debt (USD m)',
                'Type of PPI'
            ]

            for col in new_columns:
                if col not in df.columns:
                    df[col] = None

            # Map the necessary columns
            df['Transaction Country (PPI)'] = df['Transaction Country'].map(country_to_ppi)
            df['Transaction Region (PPI)'] = df['Transaction Country'].map(country_to_region_ppi)
            df['IDA Status'] = df['Transaction Country'].map(country_to_ida_status)

            # Handle the 'Transaction Subsector' and 'Transaction Sector' mapping using the 'Tranches' sheet
            tranches_df = sheet_data['Tranches']
            if not tranches_df['Realfin INFRA Transaction ID'].is_unique:
                tranches_df = tranches_df.drop_duplicates(subset='Realfin INFRA Transaction ID')

            # Correctly populate the 'Transaction Sector' column
            df['Transaction Sector'] = df['Realfin INFRA Transaction ID'].map(
                tranches_df.set_index('Realfin INFRA Transaction ID')['Transaction Sector']
            )

            # Also map 'Transaction Subsector' from 'Tranches' sheet
            df['Transaction Subsector'] = df['Realfin INFRA Transaction ID'].map(
                tranches_df.set_index('Realfin INFRA Transaction ID')['Transaction Subsector']
            )

            # Map and concatenate PPI values
            df['Transaction Sector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 0) if pd.notnull(x) else None)
            df['Transaction Subsector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 1) if pd.notnull(x) else None)
            df['Transaction Segment (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 2) if pd.notnull(x) else None)

            # Call the determine_type_of_ppi function to populate 'Type of PPI'
            df = determine_type_of_ppi(df)

            # Populate 'Commercial Bank Debt (USD m)' from 'Tranche_Participants' sheet
            if 'Tranche_Participants' in sheet_data:
                tranche_participants_df = sheet_data['Tranche_Participants']
                commercial_bank_debt = (
                    tranche_participants_df[tranche_participants_df['Participant Company Type'] == 'Commercial Bank']
                    .groupby('Realfin INFRA Transaction ID')['Participant Tranche Underwriting (USD m)']
                    .sum()
                )
                df['Commercial Bank Debt (USD m)'] = df['Realfin INFRA Transaction ID'].map(commercial_bank_debt)

                # Functionality 1: Populate 'Multilateral Debt (USD m)'
                multilateral_debt = (
                    tranche_participants_df[tranche_participants_df['Participant Company Type'] == 'Multilateral']
                    .groupby('Realfin INFRA Transaction ID')['Participant Tranche Underwriting (USD m)']
                    .sum()
                )
                df['Multilateral Debt (USD m)'] = df['Realfin INFRA Transaction ID'].map(multilateral_debt)

                # Functionality 2: Populate 'Multilateral Involvement'
                multilateral_involvement = (
                    tranche_participants_df.groupby('Realfin INFRA Transaction ID')['Participant Company Type']
                    .apply(lambda x: 'Yes' if 'Multilateral' in x.values else 'No')
                )
                df['Multilateral Involvement'] = df['Realfin INFRA Transaction ID'].map(multilateral_involvement).fillna('No')

                # Functionality 3: Populate 'Bilateral Involvement'
                bilateral_conditions = tranche_participants_df['Participant Company Type'].isin(['Development Bank', 'Export Credit Agency'])
                bilateral_involvement = (
                    tranche_participants_df[bilateral_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Company Type']
                    .apply(lambda x: 'Yes' if not x.empty else 'No')
                )
                df['Bilateral Involvement'] = df['Realfin INFRA Transaction ID'].map(bilateral_involvement).fillna('No')

                # Functionality 4: Populate 'Bilateral Debt (USD m)'
                bilateral_debt = (
                    tranche_participants_df[bilateral_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Tranche Underwriting (USD m)']
                    .sum()
                )
                df['Bilateral Debt (USD m)'] = df['Realfin INFRA Transaction ID'].map(bilateral_debt)

                # Functionality 5: Populate 'Institutional Involvement'
                institutional_conditions = tranche_participants_df['Participant Company Type'].isin([
                    "Pension Fund (Public)", "Pension Fund (Private)", "Superannuation Fund",
                    "Insurance Company", "Insurance Company (Life)", "Insurance (Non-Life)",
                    "Sovereign Wealth Fund", "Endowment", "Foundation", "Institutional - Other",
                    "Fund Manager", "Asset Manager", "Private Equity", "Fund"
                ])
                institutional_involvement = (
                    tranche_participants_df[institutional_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Company Type']
                    .apply(lambda x: 'Yes' if not x.empty else 'No')
                )
                df['Institutional Involvement'] = df['Realfin INFRA Transaction ID'].map(institutional_involvement).fillna('No')

                # Functionality 6: Populate 'Institutional Debt (USD m)'
                institutional_debt = (
                    tranche_participants_df[institutional_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Tranche Underwriting (USD m)']
                    .sum()
                )
                df['Institutional Debt (USD m)'] = df['Realfin INFRA Transaction ID'].map(institutional_debt)

                # Functionality 7: Populate 'Public Involvement'
                public_conditions = tranche_participants_df['Participant Company Type'].isin([
                    "State Bank", "State-Owned Enterprise", "Government (Local)", "Government (National)"
                ])
                public_involvement = (
                    tranche_participants_df[public_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Company Type']
                    .apply(lambda x: 'Yes' if not x.empty else 'No')
                )
                df['Public Involvement'] = df['Realfin INFRA Transaction ID'].map(public_involvement).fillna('No')

                # Functionality 8: Populate 'Public Debt (USD m)'
                public_debt = (
                    tranche_participants_df[public_conditions]
                    .groupby('Realfin INFRA Transaction ID')['Participant Tranche Underwriting (USD m)']
                    .sum()
                )
                df['Public Debt (USD m)'] = df['Realfin INFRA Transaction ID'].map(public_debt)

                # Functionality 9: Populate 'International Involvement'
                international_involvement = (
                    tranche_participants_df.apply(
                        lambda row: 'No' if row['Participant Domicile(Country)'] == row['Transaction Country'] else 'Yes',
                        axis=1
                    )
                    .groupby(tranche_participants_df['Realfin INFRA Transaction ID'])
                    .first()
                )
                df['International Involvement'] = df['Realfin INFRA Transaction ID'].map(international_involvement)

            # Ensure the correct order of the new columns
            df = df[original_columns + new_columns]

            sheet_data['Transactions_Infrastructure_GIH'] = df

        # Processing 'Tranches' sheet
        if 'Tranches' in sheet_data:
            df = sheet_data['Tranches']
            df['Transaction Country (PPI)'] = df['Transaction Country'].map(country_to_ppi)
            df['Transaction Region (PPI)'] = df['Transaction Country'].map(country_to_region_ppi)
            df['IDA Status'] = df['Transaction Country'].map(country_to_ida_status)
            
            df['Transaction Sector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 0) if pd.notnull(x) else None)
            df['Transaction Subsector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 1) if pd.notnull(x) else None)
            df['Transaction Segment (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 2) if pd.notnull(x) else None)

            df['Transaction Sector (PPI) Split'], df['Transaction Subsector (PPI) Split'], df['Transaction Segment (PPI) Split'] = zip(*df['Transaction Subsector'].map(apply_split_mappings))

            sheet_data['Tranches'] = df

        # Processing 'Tranche_Participants' sheet
        if 'Tranche_Participants' in sheet_data:
            df = sheet_data['Tranche_Participants']
            df['Transaction Country (PPI)'] = df['Transaction Country'].map(country_to_ppi)
            df['Transaction Region (PPI)'] = df['Transaction Country'].map(country_to_region_ppi)
            df['IDA Status'] = df['Transaction Country'].map(country_to_ida_status)

            df['Transaction Sector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 0) if pd.notnull(x) else None)
            df['Transaction Subsector (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 1) if pd.notnull(x) else None)
            df['Transaction Segment (PPI)'] = df['Transaction Subsector'].apply(lambda x: map_and_concatenate_values(x, 2) if pd.notnull(x) else None)

            sheet_data['Tranche_Participants'] = df

    # Get the current time in London, UK timezone
    london_tz = pytz.timezone('Europe/London')
    current_time = datetime.now(london_tz).strftime("_%d%m%y_%H%M")

    # Create a new file name with date and time suffix
    destination_file_name = f"GIH{current_time}.xlsx"
    destination_path = os.path.join(tempfile.gettempdir(), destination_file_name)

    # Write all the sheets back to a new Excel file and autofit columns
    with pd.ExcelWriter(destination_path, engine='openpyxl') as writer:
        for sheet_name, df in sheet_data.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            autofit_columns(worksheet)  # Autofit columns for each sheet

    return destination_path

# Streamlit app
st.title('Curating GIH')

uploaded_file = st.file_uploader("Choose a source file", type=["xlsx"])

if uploaded_file is not None:
    start_time = time.time()  # Start the timer once a file is uploaded
    
    # Save the uploaded file to a temporary directory
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_file.write(uploaded_file.getbuffer())
    temp_file_path = temp_file.name
    temp_file.close()  # Ensure file is closed before processing

    destination_path = None  # Initialize destination_path

    try:
        with st.spinner("Processing the file..."):
            destination_path = create_destination_file(temp_file_path, start_time)
        st.success("File processed successfully!")

        # Provide a download button for the processed file
        with open(destination_path, "rb") as file:
            st.download_button(
                label="Download Processed File",
                data=file,
                file_name=os.path.basename(destination_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"An error occurred: {e}")

    finally:
        # Clean up temporary files
        try:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
        except PermissionError:
            st.warning("Temporary file could not be deleted immediately, please try again later.")
        if destination_path and os.path.exists(destination_path):
            os.remove(destination_path)

else:
    st.info("Please upload an Excel file to start processing.")