"""
CSW Savings Calculator - Streamlit Web App
Converted from Excel workbook
"""

import streamlit as st
import pandas as pd
import numpy as np

# ============================================================================
# CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="CSW Savings Calculator",
    page_icon="🏢",
    layout="wide"
)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1

# ============================================================================
# DATA: WEATHER LOOKUP TABLE (from Weather Information sheet)
# ============================================================================

# NOTE: This is a partial list - the full Excel has 874 locations
# You'll need to add all locations from the Weather Information sheet
WEATHER_DATA = {
    'Alabama - Anniston': {'HDD': 2585, 'CDD': 1713},
    'Alabama - Auburn': {'HDD': 2688, 'CDD': 1477},
    'Alabama - Birmingham': {'HDD': 2551, 'CDD': 1667},
    'Alabama - Dothan': {'HDD': 2052, 'CDD': 2097},
    'Alabama - Huntsville': {'HDD': 3385, 'CDD': 1576},
    'Alabama - Mobile': {'HDD': 1667, 'CDD': 2448},
    'Alabama - Montgomery': {'HDD': 2191, 'CDD': 2095},
    'Alaska - Anchorage': {'HDD': 10447, 'CDD': 7},
    'Alaska - Fairbanks': {'HDD': 13969, 'CDD': 92},
    'Alaska - Juneau': {'HDD': 8966, 'CDD': 0},
    'Arizona - Flagstaff': {'HDD': 7152, 'CDD': 232},
    'Arizona - Phoenix': {'HDD': 1439, 'CDD': 2974},
    'Arizona - Tucson': {'HDD': 1846, 'CDD': 2639},
    'Arkansas - Fort Smith': {'HDD': 3397, 'CDD': 1933},
    'Arkansas - Little Rock': {'HDD': 3254, 'CDD': 1886},
    'California - Bakersfield': {'HDD': 2122, 'CDD': 1617},
    'California - Fresno': {'HDD': 2650, 'CDD': 1326},
    'California - Los Angeles': {'HDD': 1349, 'CDD': 674},
    'California - Oakland': {'HDD': 2870, 'CDD': 188},
    'California - Sacramento': {'HDD': 2502, 'CDD': 1285},
    'California - San Diego': {'HDD': 1305, 'CDD': 772},
    'California - San Francisco': {'HDD': 2901, 'CDD': 119},
    'Colorado - Colorado Springs': {'HDD': 6423, 'CDD': 471},
    'Colorado - Denver': {'HDD': 5792, 'CDD': 734},
    'Colorado - Grand Junction': {'HDD': 5641, 'CDD': 936},
    'Connecticut - Bridgeport': {'HDD': 5617, 'CDD': 689},
    'Connecticut - Hartford': {'HDD': 6171, 'CDD': 667},
    'Delaware - Wilmington': {'HDD': 4930, 'CDD': 1093},
    'District of Columbia - Washington': {'HDD': 4211, 'CDD': 1415},
    'Florida - Jacksonville': {'HDD': 1293, 'CDD': 2553},
    'Florida - Key West': {'HDD': 60, 'CDD': 4426},
    'Florida - Miami': {'HDD': 161, 'CDD': 4038},
    'Florida - Orlando': {'HDD': 667, 'CDD': 3116},
    'Florida - Tampa': {'HDD': 683, 'CDD': 3000},
    'Georgia - Atlanta': {'HDD': 2961, 'CDD': 1667},
    'Georgia - Savannah': {'HDD': 1819, 'CDD': 2253},
    'Hawaii - Honolulu': {'HDD': 0, 'CDD': 4336},
    'Idaho - Boise': {'HDD': 5809, 'CDD': 838},
    'Illinois - Chicago': {'HDD': 6536, 'CDD': 782},
    'Illinois - Peoria': {'HDD': 6025, 'CDD': 950},
    'Illinois - Springfield': {'HDD': 5619, 'CDD': 1108},
    'Indiana - Indianapolis': {'HDD': 5699, 'CDD': 1042},
    'Indiana - Fort Wayne': {'HDD': 6205, 'CDD': 761},
    'Iowa - Des Moines': {'HDD': 6710, 'CDD': 923},
    'Iowa - Sioux City': {'HDD': 7239, 'CDD': 860},
    'Kansas - Dodge City': {'HDD': 5046, 'CDD': 1428},
    'Kansas - Topeka': {'HDD': 5229, 'CDD': 1403},
    'Kansas - Wichita': {'HDD': 4687, 'CDD': 1590},
    'Kentucky - Lexington': {'HDD': 4683, 'CDD': 1193},
    'Kentucky - Louisville': {'HDD': 4660, 'CDD': 1385},
    'Louisiana - Baton Rouge': {'HDD': 1617, 'CDD': 2689},
    'Louisiana - New Orleans': {'HDD': 1465, 'CDD': 2723},
    'Louisiana - Shreveport': {'HDD': 2184, 'CDD': 2379},
    'Maine - Caribou': {'HDD': 9554, 'CDD': 214},
    'Maine - Portland': {'HDD': 7511, 'CDD': 383},
    'Maryland - Baltimore': {'HDD': 4654, 'CDD': 1167},
    'Massachusetts - Boston': {'HDD': 5621, 'CDD': 752},
    'Massachusetts - Worcester': {'HDD': 6969, 'CDD': 488},
    'Michigan - Detroit': {'HDD': 6232, 'CDD': 712},
    'Michigan - Grand Rapids': {'HDD': 6894, 'CDD': 605},
    'Michigan - Lansing': {'HDD': 6909, 'CDD': 578},
    'Minnesota - Duluth': {'HDD': 9818, 'CDD': 234},
    'Minnesota - Minneapolis': {'HDD': 7876, 'CDD': 699},
    'Minnesota - Rochester': {'HDD': 8295, 'CDD': 559},
    'Mississippi - Jackson': {'HDD': 2239, 'CDD': 2239},
    'Mississippi - Meridian': {'HDD': 2289, 'CDD': 2201},
    'Missouri - Columbia': {'HDD': 5046, 'CDD': 1379},
    'Missouri - Kansas City': {'HDD': 5336, 'CDD': 1433},
    'Missouri - St Louis': {'HDD': 4900, 'CDD': 1555},
    'Missouri - Springfield': {'HDD': 4900, 'CDD': 1384},
    'Montana - Billings': {'HDD': 7049, 'CDD': 537},
    'Montana - Great Falls': {'HDD': 7750, 'CDD': 396},
    'Montana - Helena': {'HDD': 8129, 'CDD': 260},
    'Montana - Missoula': {'HDD': 8125, 'CDD': 240},
    'Nebraska - Grand Island': {'HDD': 6530, 'CDD': 1044},
    'Nebraska - Lincoln': {'HDD': 6150, 'CDD': 1168},
    'Nebraska - Norfolk': {'HDD': 7005, 'CDD': 907},
    'Nebraska - North Platte': {'HDD': 6684, 'CDD': 867},
    'Nebraska - Omaha': {'HDD': 6612, 'CDD': 1050},
    'Nebraska - Scottsbluff': {'HDD': 6773, 'CDD': 714},
    'Nevada - Elko': {'HDD': 7112, 'CDD': 422},
    'Nevada - Ely': {'HDD': 9501, 'CDD': 186},
    'Nevada - Las Vegas': {'HDD': 2239, 'CDD': 3008},
    'Nevada - Reno': {'HDD': 6332, 'CDD': 350},
    'Nevada - Winnemucca': {'HDD': 6761, 'CDD': 493},
    'New Hampshire - Concord': {'HDD': 7383, 'CDD': 492},
    'New Jersey - Atlantic City': {'HDD': 4812, 'CDD': 978},
    'New Jersey - Newark': {'HDD': 4589, 'CDD': 1134},
    'New Mexico - Albuquerque': {'HDD': 4348, 'CDD': 1256},
    'New Mexico - Roswell': {'HDD': 3903, 'CDD': 1536},
    'New York - Albany': {'HDD': 6875, 'CDD': 579},
    'New York - Binghamton': {'HDD': 7286, 'CDD': 461},
    'New York - Buffalo': {'HDD': 6927, 'CDD': 526},
    'New York - New York City': {'HDD': 4811, 'CDD': 1089},
    'New York - Rochester': {'HDD': 6748, 'CDD': 555},
    'New York - Syracuse': {'HDD': 6756, 'CDD': 572},
    'North Carolina - Asheville': {'HDD': 4254, 'CDD': 829},
    'North Carolina - Charlotte': {'HDD': 3191, 'CDD': 1575},
    'North Carolina - Greensboro': {'HDD': 3868, 'CDD': 1371},
    'North Carolina - Raleigh': {'HDD': 3393, 'CDD': 1481},
    'North Carolina - Wilmington': {'HDD': 2347, 'CDD': 1899},
    'North Dakota - Bismarck': {'HDD': 8851, 'CDD': 541},
    'North Dakota - Fargo': {'HDD': 9226, 'CDD': 501},
    'Ohio - Akron': {'HDD': 6037, 'CDD': 687},
    'Ohio - Cincinnati': {'HDD': 4410, 'CDD': 1230},
    'Ohio - Cleveland': {'HDD': 6351, 'CDD': 629},
    'Ohio - Columbus': {'HDD': 5660, 'CDD': 874},
    'Ohio - Dayton': {'HDD': 5622, 'CDD': 966},
    'Ohio - Toledo': {'HDD': 6494, 'CDD': 708},
    'Ohio - Youngstown': {'HDD': 6417, 'CDD': 579},
    'Oklahoma - Oklahoma City': {'HDD': 3725, 'CDD': 1924},
    'Oklahoma - Tulsa': {'HDD': 3860, 'CDD': 1945},
    'Oregon - Eugene': {'HDD': 4726, 'CDD': 253},
    'Oregon - Medford': {'HDD': 4930, 'CDD': 699},
    'Oregon - Pendleton': {'HDD': 5127, 'CDD': 741},
    'Oregon - Portland': {'HDD': 4635, 'CDD': 371},
    'Oregon - Salem': {'HDD': 4715, 'CDD': 336},
    'Pennsylvania - Allentown': {'HDD': 5827, 'CDD': 782},
    'Pennsylvania - Erie': {'HDD': 6451, 'CDD': 535},
    'Pennsylvania - Harrisburg': {'HDD': 5251, 'CDD': 1050},
    'Pennsylvania - Philadelphia': {'HDD': 4865, 'CDD': 1152},
    'Pennsylvania - Pittsburgh': {'HDD': 5987, 'CDD': 692},
    'Pennsylvania - Williamsport': {'HDD': 5934, 'CDD': 737},
    'Rhode Island - Providence': {'HDD': 5804, 'CDD': 660},
    'South Carolina - Charleston': {'HDD': 2033, 'CDD': 2219},
    'South Carolina - Columbia': {'HDD': 2484, 'CDD': 1989},
    'South Carolina - Greenville': {'HDD': 3205, 'CDD': 1524},
    'South Dakota - Huron': {'HDD': 8223, 'CDD': 715},
    'South Dakota - Rapid City': {'HDD': 7345, 'CDD': 651},
    'South Dakota - Sioux Falls': {'HDD': 7839, 'CDD': 783},
    'Tennessee - Bristol': {'HDD': 4143, 'CDD': 1012},
    'Tennessee - Chattanooga': {'HDD': 3254, 'CDD': 1580},
    'Tennessee - Knoxville': {'HDD': 3494, 'CDD': 1466},
    'Tennessee - Memphis': {'HDD': 3015, 'CDD': 2126},
    'Tennessee - Nashville': {'HDD': 3696, 'CDD': 1614},
    'Texas - Abilene': {'HDD': 2624, 'CDD': 2466},
    'Texas - Amarillo': {'HDD': 4182, 'CDD': 1445},
    'Texas - Austin': {'HDD': 1711, 'CDD': 2887},
    'Texas - Brownsville': {'HDD': 600, 'CDD': 3667},
    'Texas - Corpus Christi': {'HDD': 914, 'CDD': 3222},
    'Texas - Dallas': {'HDD': 2363, 'CDD': 2583},
    'Texas - El Paso': {'HDD': 2700, 'CDD': 1745},
    'Texas - Fort Worth': {'HDD': 2405, 'CDD': 2570},
    'Texas - Houston': {'HDD': 1434, 'CDD': 2823},
    'Texas - Lubbock': {'HDD': 3735, 'CDD': 1647},
    'Texas - Midland': {'HDD': 2591, 'CDD': 2155},
    'Texas - Port Arthur': {'HDD': 1426, 'CDD': 2695},
    'Texas - San Angelo': {'HDD': 2413, 'CDD': 2417},
    'Texas - San Antonio': {'HDD': 1546, 'CDD': 2900},
    'Texas - Victoria': {'HDD': 1159, 'CDD': 2952},
    'Texas - Waco': {'HDD': 2030, 'CDD': 2625},
    'Texas - Wichita Falls': {'HDD': 2832, 'CDD': 2405},
    'Utah - Cedar City': {'HDD': 6143, 'CDD': 617},
    'Utah - Salt Lake City': {'HDD': 5983, 'CDD': 944},
    'Vermont - Burlington': {'HDD': 8269, 'CDD': 410},
    'Virginia - Lynchburg': {'HDD': 4166, 'CDD': 1144},
    'Virginia - Norfolk': {'HDD': 3421, 'CDD': 1495},
    'Virginia - Richmond': {'HDD': 3865, 'CDD': 1399},
    'Virginia - Roanoke': {'HDD': 4150, 'CDD': 1100},
    'Washington - Olympia': {'HDD': 5236, 'CDD': 122},
    'Washington - Seattle': {'HDD': 4797, 'CDD': 183},
    'Washington - Spokane': {'HDD': 6655, 'CDD': 402},
    'Washington - Walla Walla': {'HDD': 5128, 'CDD': 805},
    'Washington - Yakima': {'HDD': 5457, 'CDD': 620},
    'West Virginia - Charleston': {'HDD': 4476, 'CDD': 1076},
    'West Virginia - Elkins': {'HDD': 5675, 'CDD': 533},
    'West Virginia - Huntington': {'HDD': 4446, 'CDD': 1169},
    'Wisconsin - Green Bay': {'HDD': 8029, 'CDD': 474},
    'Wisconsin - La Crosse': {'HDD': 7589, 'CDD': 647},
    'Wisconsin - Madison': {'HDD': 7863, 'CDD': 591},
    'Wisconsin - Milwaukee': {'HDD': 7635, 'CDD': 515},
    'Wyoming - Casper': {'HDD': 7410, 'CDD': 373},
    'Wyoming - Cheyenne': {'HDD': 7381, 'CDD': 347},
    'Wyoming - Lander': {'HDD': 7870, 'CDD': 275},
    'Wyoming - Sheridan': {'HDD': 7680, 'CDD': 470},
}

# ============================================================================
# DATA: DROPDOWN OPTIONS
# ============================================================================

HVAC_SYSTEMS = [
    'Packaged VAV with electric reheat',
    'Built-up VAV with hydronic reheat',
    'Packaged single zone',
    'PVAV with gas heat',
    'PVAV with PFP boxes'
]

HEATING_FUELS = ['Electric', 'Natural Gas', 'None']

COOLING_OPTIONS = ['Yes', 'No']

WINDOW_TYPES = ['Single pane', 'Double pane', 'Double pane, low-e']

CSW_TYPES = ['Double', 'Triple', 'Quad']

# ============================================================================
# DATA: SAVINGS LOOKUP TABLE
# ============================================================================

# This is a SIMPLIFIED version - the actual Excel has 104 rows with complex configs
# Format: {config_key: {heating, cooling, gas, ...}}
# Config key format: [ExistWin][NewWin][Size][Building][HVAC]_[Fuel][HeatFuel][Hours]

SAVINGS_LOOKUP = {
    # Single->Double, Mid Office, PVAV Electric, 8760 hours
    'SingleDoubleMidOfficePVAV_ElecElectric8760': {
        'heating_8760': 3.16,
        'cooling_8760': 3.99,
        'gas_8760': 0,
        'baseline_8760': 140.38
    },
    # Single->Double, Mid Office, PVAV Electric, 2080 hours
    'SingleDoubleMidOfficePVAV_ElecElectric2080': {
        'heating_2080': 3.35,
        'cooling_2080': 7.67,
        'gas_2080': 0,
        'baseline_2080': 130.37
    },
    # Add more lookup combinations as needed
}

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def determine_building_size(building_area, hvac_system):
    """Determine if building is Mid or Large size (from Excel M24 formula)"""
    if building_area > 30000 and hvac_system == "Built-up VAV with hydronic reheat":
        return "Large"
    else:
        return "Mid"

def get_window_code(window_type):
    """Convert window type to code used in lookup key"""
    if window_type == "Single pane":
        return "Single"
    elif window_type == "Double pane":
        return "Double"
    elif window_type == "Double pane, low-e":
        return "Double"  # Low-e treated as double for lookup
    return "Double"

def interpolate_operating_hours(operating_hours, val_8760, val_2080):
    """
    Interpolate between 2080 and 8760 hour values based on operating hours
    From Excel formula: IF($F$23<2080,val_2080,IF($F$23=8760,val_8760,(($F$23-2080)/(8760-2080)*(val_8760-val_2080)+val_2080)))
    """
    if operating_hours < 2080:
        return val_2080
    elif operating_hours >= 8760:
        return val_8760
    else:
        # Linear interpolation
        return ((operating_hours - 2080) / (8760 - 2080)) * (val_8760 - val_2080) + val_2080

def calculate_savings(inputs):
    """
    Main calculation function
    Implements Excel formulas from Office sheet
    """
    # Extract inputs
    building_area = inputs['building_area']
    csw_area = inputs['csw_area']
    operating_hours = inputs['operating_hours']
    heating_fuel = inputs['heating_fuel']
    electric_rate = inputs['electric_rate']
    gas_rate = inputs['gas_rate']
    location = inputs['location']
    existing_window = inputs['existing_window']
    csw_type = inputs['csw_type']
    hvac_system = inputs['hvac_system']
    cooling_installed = inputs['cooling_installed']
    
    # Get weather data (Excel C23, C24)
    weather = WEATHER_DATA.get(location, {'HDD': 0, 'CDD': 0})
    hdd = weather['HDD']
    cdd = weather['CDD']
    
    # Determine building size (Excel M24)
    building_size = determine_building_size(building_area, hvac_system)
    
    # Build lookup key for SAVINGS_LOOKUP table
    existing_code = get_window_code(existing_window)
    new_code = get_window_code(csw_type)
    
    # Simplified lookup key (actual Excel has more complex logic)
    lookup_key = f"{existing_code}{new_code}{building_size}OfficePVAV_Elec{heating_fuel}{operating_hours}"
    
    # Try to find exact match, otherwise use default
    if lookup_key not in SAVINGS_LOOKUP:
        # Use default Single->Double, Mid, Electric, 8760
        lookup_key = 'SingleDoubleMidOfficePVAV_ElecElectric8760'
    
    lookup_data = SAVINGS_LOOKUP.get(lookup_key, {})
    
    # Get values for interpolation
    heating_8760 = lookup_data.get('heating_8760', 3.16)
    cooling_8760 = lookup_data.get('cooling_8760', 3.99)
    heating_2080 = lookup_data.get('heating_2080', 3.35)
    cooling_2080 = lookup_data.get('cooling_2080', 7.67)
    baseline_8760 = lookup_data.get('baseline_8760', 140.38)
    baseline_2080 = lookup_data.get('baseline_2080', 130.37)
    
    # Interpolate values based on operating hours (Excel C31, C32)
    heating_savings_per_sf = interpolate_operating_hours(operating_hours, heating_8760, heating_2080)
    
    # Cooling calculation includes cooling factor if installed
    if cooling_installed == "Yes":
        cooling_factor = 1.0  # Simplified - actual Excel has more complex logic
    else:
        cooling_factor = 0.0
    
    cooling_savings_per_sf = interpolate_operating_hours(operating_hours, cooling_8760, cooling_2080) * cooling_factor
    
    # Calculate total electric savings per SF (Excel C31 + C32)
    total_electric_per_sf = heating_savings_per_sf + cooling_savings_per_sf
    
    # Calculate total electric savings (Excel F31 = (C31+C32)*F27)
    electric_savings_kwh = total_electric_per_sf * csw_area
    
    # Gas savings (Excel F33 = C33*F27)
    # Simplified - gas savings depend on heating fuel type
    if heating_fuel == "Natural Gas":
        gas_savings_per_sf = 0.59  # Placeholder from Excel
        gas_savings_therms = gas_savings_per_sf * csw_area
    else:
        gas_savings_therms = 0
    
    # Calculate cost savings (Excel C35, C36)
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    # Calculate savings intensity (Excel F35 = (F31*3.413+F33*100)/F18)
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    
    # Baseline EUI (Excel F13 - interpolated)
    baseline_eui = interpolate_operating_hours(operating_hours, baseline_8760, baseline_2080)
    
    # Estimated Window-to-Wall Ratio (Excel Z26 - simplified)
    wwr = 0.33  # Placeholder - actual calculation is complex regression
    
    return {
        'electric_savings_kwh': electric_savings_kwh,
        'gas_savings_therms': gas_savings_therms,
        'electric_cost_savings': electric_cost_savings,
        'gas_cost_savings': gas_cost_savings,
        'total_cost_savings': total_cost_savings,
        'total_savings_kbtu_sf': total_savings_kbtu_sf,
        'baseline_eui': baseline_eui,
        'wwr': wwr,
        'hdd': hdd,
        'cdd': cdd
    }

# ============================================================================
# STREAMLIT UI
# ============================================================================

st.title('🏢 Commercial Window Savings Calculator')
st.markdown('### Office Buildings')
st.markdown('---')

# Progress indicator
progress = st.session_state.step / 4
st.progress(progress)
st.write(f'Step {st.session_state.step} of 4')

# ============================================================================
# STEP 1: LOCATION
# ============================================================================

if st.session_state.step == 1:
    st.header('Step 1: Project Location')
    
    col1, col2 = st.columns(2)
    
    with col1:
        location = st.selectbox(
            'Select Project Location (City, State)',
            options=sorted(WEATHER_DATA.keys()),
            key='location'
        )
    
    if location and location in WEATHER_DATA:
        with col2:
            st.info(f"**Location HDD (Base 65):** {WEATHER_DATA[location]['HDD']}")
            st.info(f"**Location CDD (Base 65):** {WEATHER_DATA[location]['CDD']}")
    
    if st.button('Next →', type='primary'):
        st.session_state.step = 2
        st.rerun()

# ============================================================================
# STEP 2: BUILDING INFORMATION
# ============================================================================

elif st.session_state.step == 2:
    st.header('Step 2: Building Information')
    
    col1, col2 = st.columns(2)
    
    with col1:
        building_area = st.number_input(
            'Building Area (Sq.Ft.)',
            min_value=15000,
            max_value=500000,
            value=75000,
            step=1000,
            key='building_area',
            help='Minimum: 15,000 sq.ft., Maximum: 500,000 sq.ft.'
        )
        
        num_floors = st.number_input(
            'Number of Floors',
            min_value=1,
            max_value=50,
            value=5,
            key='num_floors'
        )
        
        operating_hours = st.number_input(
            'Annual Operating Hours',
            min_value=1980,
            max_value=8760,
            value=8000,
            key='operating_hours',
            help='Minimum: 1,980 hrs/yr, Maximum: 8,760 hrs/yr'
        )
    
    with col2:
        hvac_system = st.selectbox(
            'HVAC System Type',
            options=HVAC_SYSTEMS,
            key='hvac_system'
        )
        
        heating_fuel = st.selectbox(
            'Heating Fuel',
            options=HEATING_FUELS,
            key='heating_fuel'
        )
        
        cooling_installed = st.selectbox(
            'Cooling Installed?',
            options=COOLING_OPTIONS,
            key='cooling_installed'
        )
    
    # Calculate and display Window-to-Wall Ratio estimate
    if building_area and num_floors:
        # Simplified WWR estimation
        wwr = 0.33  # Placeholder - actual calculation uses regression
        st.info(f"**Est. Window-to-Wall Ratio:** {wwr:.1%}")
    
    # Validation warnings
    if building_area < 15000 or building_area > 500000:
        st.warning('⚠️ Building area outside recommended range (15,000 - 500,000 sq.ft.)')
    
    if operating_hours < 1980 or operating_hours > 8760:
        st.warning('⚠️ Operating hours outside valid range (1,980 - 8,760 hrs/yr)')
    
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button('← Back'):
            st.session_state.step = 1
            st.rerun()
    with col_next:
        if st.button('Next →', type='primary'):
            st.session_state.step = 3
            st.rerun()

# ============================================================================
# STEP 3: WINDOW INFORMATION
# ============================================================================

elif st.session_state.step == 3:
    st.header('Step 3: Window & Cost Information')
    
    col1, col2 = st.columns(2)
    
    with col1:
        existing_window = st.selectbox(
            'Type of Existing Window',
            options=WINDOW_TYPES,
            key='existing_window',
            help='Select the current window type in your building'
        )
        
        csw_type = st.selectbox(
            'Type of CSW Analyzed',
            options=CSW_TYPES,
            key='csw_type',
            help='Commercial storm window type to be installed'
        )
        
        csw_area = st.number_input(
            'Sq.ft. of CSW Installed',
            min_value=0,
            max_value=int(st.session_state.get('building_area', 75000) * 0.5),
            value=12000,
            step=100,
            key='csw_area',
            help='Total square footage of commercial storm windows to install'
        )
    
    with col2:
        electric_rate = st.number_input(
            'Electric Rate ($/kWh)',
            min_value=0.01,
            max_value=1.0,
            value=0.12,
            step=0.01,
            format='%.3f',
            key='electric_rate',
            help='Your current electricity rate'
        )
        
        gas_rate = st.number_input(
            'Natural Gas Rate ($/therm)',
            min_value=0.01,
            max_value=10.0,
            value=0.80,
            step=0.05,
            format='%.2f',
            key='gas_rate',
            help='Your current natural gas rate'
        )
    
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button('← Back'):
            st.session_state.step = 2
            st.rerun()
    with col_next:
        if st.button('Calculate Savings →', type='primary'):
            st.session_state.step = 4
            st.rerun()

# ============================================================================
# STEP 4: RESULTS
# ============================================================================

elif st.session_state.step == 4:
    st.header('💡 Your Energy Savings Results')
    
    # Gather all inputs
    inputs = {
        'location': st.session_state.get('location', list(WEATHER_DATA.keys())[0]),
        'building_area': st.session_state.get('building_area', 75000),
        'num_floors': st.session_state.get('num_floors', 5),
        'operating_hours': st.session_state.get('operating_hours', 8000),
        'hvac_system': st.session_state.get('hvac_system', HVAC_SYSTEMS[0]),
        'heating_fuel': st.session_state.get('heating_fuel', 'Electric'),
        'cooling_installed': st.session_state.get('cooling_installed', 'Yes'),
        'existing_window': st.session_state.get('existing_window', 'Single pane'),
        'csw_type': st.session_state.get('csw_type', 'Double'),
        'csw_area': st.session_state.get('csw_area', 12000),
        'electric_rate': st.session_state.get('electric_rate', 0.12),
        'gas_rate': st.session_state.get('gas_rate', 0.80)
    }
    
    # Calculate savings
    results = calculate_savings(inputs)
    
    st.success('✅ Calculation Complete!')
    
    # Display intermediary calculations
    with st.expander('📊 Project Details', expanded=True):
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric('Est. Window-to-Wall Ratio', f"{results['wwr']:.1%}")
        with col2:
            st.metric('Location HDD (Base 65)', f"{results['hdd']:,}")
        with col3:
            st.metric('Location CDD (Base 65)', f"{results['cdd']:,}")
    
    st.markdown('---')
    
    # Main results
    st.markdown('### 💰 Annual Energy Savings')
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            label='⚡ Electric Energy Savings',
            value=f'{results["electric_savings_kwh"]:,.0f} kWh/yr'
        )
        st.metric(
            label='🔥 Gas Energy Savings',
            value=f'{results["gas_savings_therms"]:,.0f} therms/yr'
        )
    
    with col2:
        st.metric(
            label='💵 Electric Cost Savings',
            value=f'${results["electric_cost_savings"]:,.2f}/yr'
        )
        st.metric(
            label='💵 Gas Cost Savings',
            value=f'${results["gas_cost_savings"]:,.2f}/yr'
        )
    
    with col3:
        st.metric(
            label='📊 Total Savings',
            value=f'{results["total_savings_kbtu_sf"]:.2f} kBtu/SF-yr'
        )
        st.metric(
            label='💰 Total Cost Savings',
            value=f'${results["total_cost_savings"]:,.2f}/yr',
            delta=f'${results["total_cost_savings"]/12:,.2f}/month'
        )
    
    st.markdown('---')
    
    st.metric(
        label='📈 Baseline Energy Use Intensity',
        value=f'{results["baseline_eui"]:.2f} kBtu/SF-yr'
    )
    
    # Summary table
    st.markdown('### 📋 Summary Report')
    summary_df = pd.DataFrame({
        'Metric': [
            'Electric Energy Savings',
            'Gas Energy Savings',
            'Electric Cost Savings',
            'Gas Cost Savings',
            'Total Savings (Intensity)',
            'Total Cost Savings',
            'Baseline Energy Use Intensity'
        ],
        'Value': [
            f'{results["electric_savings_kwh"]:,.0f} kWh/yr',
            f'{results["gas_savings_therms"]:,.0f} therms/yr',
            f'${results["electric_cost_savings"]:,.2f}/yr',
            f'${results["gas_cost_savings"]:,.2f}/yr',
            f'{results["total_savings_kbtu_sf"]:.2f} kBtu/SF-yr',
            f'${results["total_cost_savings"]:,.2f}/yr',
            f'{results["baseline_eui"]:.2f} kBtu/SF-yr'
        ]
    })
    
    st.dataframe(summary_df, use_container_width=True, hide_index=True)
    
    # Download button for results
    csv = summary_df.to_csv(index=False)
    st.download_button(
        label='📥 Download Results (CSV)',
        data=csv,
        file_name='csw_savings_results.csv',
        mime='text/csv'
    )
    
    st.markdown('---')
    
    if st.button('← Start Over', type='secondary'):
        st.session_state.step = 1
        st.rerun()

# ============================================================================
# SIDEBAR: INPUT SUMMARY
# ============================================================================

with st.sidebar:
    st.markdown('### 📝 Input Summary')
    
    if st.session_state.step > 1:
        st.markdown(f"**Location:** {st.session_state.get('location', 'Not set')}")
    
    if st.session_state.step > 2:
        st.markdown(f"**Building Area:** {st.session_state.get('building_area', 0):,} sq.ft.")
        st.markdown(f"**Floors:** {st.session_state.get('num_floors', 0)}")
        st.markdown(f"**Operating Hours:** {st.session_state.get('operating_hours', 0):,} hrs/yr")
        st.markdown(f"**HVAC System:** {st.session_state.get('hvac_system', 'Not set')}")
        st.markdown(f"**Heating Fuel:** {st.session_state.get('heating_fuel', 'Not set')}")
        st.markdown(f"**Cooling:** {st.session_state.get('cooling_installed', 'Not set')}")
    
    if st.session_state.step > 3:
        st.markdown(f"**Existing Window:** {st.session_state.get('existing_window', 'Not set')}")
        st.markdown(f"**CSW Type:** {st.session_state.get('csw_type', 'Not set')}")
        st.markdown(f"**CSW Area:** {st.session_state.get('csw_area', 0):,} sq.ft.")
    
    st.markdown('---')
    st.markdown('#### ⚠️ Note')
    st.markdown('This calculator uses simplified savings calculations. For more accurate results with complete lookup tables, the Excel data needs to be fully extracted.')
    
    st.markdown('---')
    st.markdown('#### 🔧 Next Steps')
    st.markdown("""
    To improve this calculator:
    1. Extract all 874 weather locations
    2. Extract all 104 savings lookup rows
    3. Implement full regression formulas
    4. Add validation rules
    5. Add lead capture form
    """)

# ============================================================================
# FOOTER
# ============================================================================

st.markdown('---')
st.markdown(
    '<div style="text-align: center; color: gray; font-size: 0.8em;">'
    'CSW Savings Calculator | Office Buildings Only'
    '</div>',
    unsafe_allow_html=True
)