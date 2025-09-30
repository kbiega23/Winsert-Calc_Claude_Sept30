"""
CSW Savings Calculator - Streamlit Web App
Converted from Excel workbook - UPDATED VERSION
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
# DATA: WEATHER LOOKUP TABLE - ORGANIZED BY STATE
# ============================================================================

# NOTE: Please verify these HDD/CDD values match your Excel exactly
# Update any that are incorrect
WEATHER_DATA_BY_STATE = {
    'Alabama': {
        'Anniston': {'HDD': 2585, 'CDD': 1713},
        'Auburn': {'HDD': 2688, 'CDD': 1477},
        'Birmingham': {'HDD': 2551, 'CDD': 1667},
        'Dothan': {'HDD': 2052, 'CDD': 2097},
        'Huntsville': {'HDD': 3385, 'CDD': 1576},
        'Mobile': {'HDD': 1667, 'CDD': 2448},
        'Montgomery': {'HDD': 2191, 'CDD': 2095},
        'Tuscaloosa': {'HDD': 2494, 'CDD': 1816},
    },
    'Alaska': {
        'Anchorage': {'HDD': 10447, 'CDD': 7},
        'Annette': {'HDD': 6447, 'CDD': 0},
        'Barrow': {'HDD': 19990, 'CDD': 0},
        'Bethel': {'HDD': 13899, 'CDD': 0},
        'Fairbanks': {'HDD': 13969, 'CDD': 92},
        'Juneau': {'HDD': 8966, 'CDD': 0},
    },
    'Arizona': {
        'Flagstaff': {'HDD': 7152, 'CDD': 232},
        'Phoenix': {'HDD': 1439, 'CDD': 2974},
        'Prescott': {'HDD': 5053, 'CDD': 556},
        'Tucson': {'HDD': 1846, 'CDD': 2639},
        'Winslow': {'HDD': 4626, 'CDD': 1097},
        'Yuma': {'HDD': 1011, 'CDD': 3837},
    },
    'Arkansas': {
        'Fort Smith': {'HDD': 3397, 'CDD': 1933},
        'Little Rock': {'HDD': 3254, 'CDD': 1886},
    },
    'California': {
        'Bakersfield': {'HDD': 2122, 'CDD': 1617},
        'Fresno': {'HDD': 2650, 'CDD': 1326},
        'Los Angeles': {'HDD': 1349, 'CDD': 674},
        'Oakland': {'HDD': 2870, 'CDD': 188},
        'Sacramento': {'HDD': 2502, 'CDD': 1285},
        'San Diego': {'HDD': 1305, 'CDD': 772},
        'San Francisco': {'HDD': 2901, 'CDD': 119},
        'Santa Barbara': {'HDD': 2298, 'CDD': 198},
    },
    'Colorado': {
        'Colorado Springs': {'HDD': 6423, 'CDD': 471},
        'Denver': {'HDD': 5792, 'CDD': 734},
        'Grand Junction': {'HDD': 5641, 'CDD': 936},
        'Pueblo': {'HDD': 5462, 'CDD': 994},
    },
    'Connecticut': {
        'Bridgeport': {'HDD': 5617, 'CDD': 689},
        'Hartford': {'HDD': 6171, 'CDD': 667},
    },
    'Delaware': {
        'Wilmington': {'HDD': 4930, 'CDD': 1093},
    },
    'District of Columbia': {
        'Washington': {'HDD': 4211, 'CDD': 1415},
    },
    'Florida': {
        'Jacksonville': {'HDD': 1293, 'CDD': 2553},
        'Key West': {'HDD': 60, 'CDD': 4426},
        'Miami': {'HDD': 161, 'CDD': 4038},
        'Orlando': {'HDD': 667, 'CDD': 3116},
        'Pensacola': {'HDD': 1523, 'CDD': 2501},
        'Tallahassee': {'HDD': 1637, 'CDD': 2566},
        'Tampa': {'HDD': 683, 'CDD': 3000},
    },
    'Georgia': {
        'Athens': {'HDD': 2951, 'CDD': 1667},
        'Atlanta': {'HDD': 2961, 'CDD': 1667},
        'Augusta': {'HDD': 2426, 'CDD': 2010},
        'Columbus': {'HDD': 2383, 'CDD': 2060},
        'Macon': {'HDD': 2136, 'CDD': 2135},
        'Savannah': {'HDD': 1819, 'CDD': 2253},
    },
    'Hawaii': {
        'Hilo': {'HDD': 0, 'CDD': 3935},
        'Honolulu': {'HDD': 0, 'CDD': 4336},
    },
    'Idaho': {
        'Boise': {'HDD': 5809, 'CDD': 838},
        'Idaho Falls': {'HDD': 7867, 'CDD': 336},
        'Pocatello': {'HDD': 7033, 'CDD': 530},
    },
    'Illinois': {
        'Chicago': {'HDD': 6536, 'CDD': 782},
        'Moline': {'HDD': 6438, 'CDD': 898},
        'Peoria': {'HDD': 6025, 'CDD': 950},
        'Rockford': {'HDD': 6830, 'CDD': 714},
        'Springfield': {'HDD': 5619, 'CDD': 1108},
    },
    'Indiana': {
        'Evansville': {'HDD': 4435, 'CDD': 1373},
        'Fort Wayne': {'HDD': 6205, 'CDD': 761},
        'Indianapolis': {'HDD': 5699, 'CDD': 1042},
        'South Bend': {'HDD': 6439, 'CDD': 697},
    },
    'Iowa': {
        'Des Moines': {'HDD': 6710, 'CDD': 923},
        'Sioux City': {'HDD': 7239, 'CDD': 860},
        'Waterloo': {'HDD': 7320, 'CDD': 800},
    },
    'Kansas': {
        'Dodge City': {'HDD': 5046, 'CDD': 1428},
        'Topeka': {'HDD': 5229, 'CDD': 1403},
        'Wichita': {'HDD': 4687, 'CDD': 1590},
    },
    'Kentucky': {
        'Lexington': {'HDD': 4683, 'CDD': 1193},
        'Louisville': {'HDD': 4660, 'CDD': 1385},
    },
    'Louisiana': {
        'Baton Rouge': {'HDD': 1617, 'CDD': 2689},
        'New Orleans': {'HDD': 1465, 'CDD': 2723},
        'Shreveport': {'HDD': 2184, 'CDD': 2379},
    },
    'Maine': {
        'Caribou': {'HDD': 9554, 'CDD': 214},
        'Portland': {'HDD': 7511, 'CDD': 383},
    },
    'Maryland': {
        'Baltimore': {'HDD': 4654, 'CDD': 1167},
    },
    'Massachusetts': {
        'Boston': {'HDD': 5621, 'CDD': 752},
        'Worcester': {'HDD': 6969, 'CDD': 488},
    },
    'Michigan': {
        'Detroit': {'HDD': 6232, 'CDD': 712},
        'Grand Rapids': {'HDD': 6894, 'CDD': 605},
        'Lansing': {'HDD': 6909, 'CDD': 578},
    },
    'Minnesota': {
        'Duluth': {'HDD': 9818, 'CDD': 234},
        'Minneapolis': {'HDD': 7876, 'CDD': 699},
        'Rochester': {'HDD': 8295, 'CDD': 559},
    },
    'Mississippi': {
        'Jackson': {'HDD': 2239, 'CDD': 2239},
        'Meridian': {'HDD': 2289, 'CDD': 2201},
    },
    'Missouri': {
        'Columbia': {'HDD': 5046, 'CDD': 1379},
        'Kansas City': {'HDD': 5336, 'CDD': 1433},
        'St Louis': {'HDD': 4900, 'CDD': 1555},
        'Springfield': {'HDD': 4900, 'CDD': 1384},
    },
    'Montana': {
        'Billings': {'HDD': 7049, 'CDD': 537},
        'Great Falls': {'HDD': 7750, 'CDD': 396},
        'Helena': {'HDD': 8129, 'CDD': 260},
        'Missoula': {'HDD': 8125, 'CDD': 240},
    },
    'Nebraska': {
        'Grand Island': {'HDD': 6530, 'CDD': 1044},
        'Lincoln': {'HDD': 6150, 'CDD': 1168},
        'Omaha': {'HDD': 6612, 'CDD': 1050},
    },
    'Nevada': {
        'Elko': {'HDD': 7112, 'CDD': 422},
        'Las Vegas': {'HDD': 2239, 'CDD': 3008},
        'Reno': {'HDD': 6332, 'CDD': 350},
    },
    'New Hampshire': {
        'Concord': {'HDD': 7383, 'CDD': 492},
    },
    'New Jersey': {
        'Atlantic City': {'HDD': 4812, 'CDD': 978},
        'Newark': {'HDD': 4589, 'CDD': 1134},
    },
    'New Mexico': {
        'Albuquerque': {'HDD': 4348, 'CDD': 1256},
        'Roswell': {'HDD': 3903, 'CDD': 1536},
    },
    'New York': {
        'Albany': {'HDD': 6875, 'CDD': 579},
        'Binghamton': {'HDD': 7286, 'CDD': 461},
        'Buffalo': {'HDD': 6927, 'CDD': 526},
        'New York City': {'HDD': 4811, 'CDD': 1089},
        'Rochester': {'HDD': 6748, 'CDD': 555},
        'Syracuse': {'HDD': 6756, 'CDD': 572},
    },
    'North Carolina': {
        'Asheville': {'HDD': 4254, 'CDD': 829},
        'Charlotte': {'HDD': 3191, 'CDD': 1575},
        'Greensboro': {'HDD': 3868, 'CDD': 1371},
        'Raleigh': {'HDD': 3393, 'CDD': 1481},
        'Wilmington': {'HDD': 2347, 'CDD': 1899},
    },
    'North Dakota': {
        'Bismarck': {'HDD': 8851, 'CDD': 541},
        'Fargo': {'HDD': 9226, 'CDD': 501},
    },
    'Ohio': {
        'Akron': {'HDD': 6037, 'CDD': 687},
        'Cincinnati': {'HDD': 4410, 'CDD': 1230},
        'Cleveland': {'HDD': 6351, 'CDD': 629},
        'Columbus': {'HDD': 5660, 'CDD': 874},
        'Dayton': {'HDD': 5622, 'CDD': 966},
        'Toledo': {'HDD': 6494, 'CDD': 708},
    },
    'Oklahoma': {
        'Oklahoma City': {'HDD': 3725, 'CDD': 1924},
        'Tulsa': {'HDD': 3860, 'CDD': 1945},
    },
    'Oregon': {
        'Eugene': {'HDD': 4726, 'CDD': 253},
        'Medford': {'HDD': 4930, 'CDD': 699},
        'Portland': {'HDD': 4635, 'CDD': 371},
        'Salem': {'HDD': 4715, 'CDD': 336},
    },
    'Pennsylvania': {
        'Allentown': {'HDD': 5827, 'CDD': 782},
        'Erie': {'HDD': 6451, 'CDD': 535},
        'Harrisburg': {'HDD': 5251, 'CDD': 1050},
        'Philadelphia': {'HDD': 4865, 'CDD': 1152},
        'Pittsburgh': {'HDD': 5987, 'CDD': 692},
    },
    'Rhode Island': {
        'Providence': {'HDD': 5804, 'CDD': 660},
    },
    'South Carolina': {
        'Charleston': {'HDD': 2033, 'CDD': 2219},
        'Columbia': {'HDD': 2484, 'CDD': 1989},
        'Greenville': {'HDD': 3205, 'CDD': 1524},
    },
    'South Dakota': {
        'Huron': {'HDD': 8223, 'CDD': 715},
        'Rapid City': {'HDD': 7345, 'CDD': 651},
        'Sioux Falls': {'HDD': 7839, 'CDD': 783},
    },
    'Tennessee': {
        'Chattanooga': {'HDD': 3254, 'CDD': 1580},
        'Knoxville': {'HDD': 3494, 'CDD': 1466},
        'Memphis': {'HDD': 3015, 'CDD': 2126},
        'Nashville': {'HDD': 3696, 'CDD': 1614},
    },
    'Texas': {
        'Abilene': {'HDD': 2624, 'CDD': 2466},
        'Amarillo': {'HDD': 4182, 'CDD': 1445},
        'Austin': {'HDD': 1711, 'CDD': 2887},
        'Dallas': {'HDD': 2363, 'CDD': 2583},
        'El Paso': {'HDD': 2700, 'CDD': 1745},
        'Fort Worth': {'HDD': 2405, 'CDD': 2570},
        'Houston': {'HDD': 1434, 'CDD': 2823},
        'Lubbock': {'HDD': 3735, 'CDD': 1647},
        'San Antonio': {'HDD': 1546, 'CDD': 2900},
    },
    'Utah': {
        'Salt Lake City': {'HDD': 5983, 'CDD': 944},
    },
    'Vermont': {
        'Burlington': {'HDD': 8269, 'CDD': 410},
    },
    'Virginia': {
        'Lynchburg': {'HDD': 4166, 'CDD': 1144},
        'Norfolk': {'HDD': 3421, 'CDD': 1495},
        'Richmond': {'HDD': 3865, 'CDD': 1399},
        'Roanoke': {'HDD': 4150, 'CDD': 1100},
    },
    'Washington': {
        'Olympia': {'HDD': 5236, 'CDD': 122},
        'Seattle': {'HDD': 4797, 'CDD': 183},
        'Spokane': {'HDD': 6655, 'CDD': 402},
        'Yakima': {'HDD': 5457, 'CDD': 620},
    },
    'West Virginia': {
        'Charleston': {'HDD': 4476, 'CDD': 1076},
        'Huntington': {'HDD': 4446, 'CDD': 1169},
    },
    'Wisconsin': {
        'Green Bay': {'HDD': 8029, 'CDD': 474},
        'Madison': {'HDD': 7863, 'CDD': 591},
        'Milwaukee': {'HDD': 7635, 'CDD': 515},
    },
    'Wyoming': {
        'Casper': {'HDD': 7410, 'CDD': 373},
        'Cheyenne': {'HDD': 7381, 'CDD': 347},
        'Sheridan': {'HDD': 7680, 'CDD': 470},
    },
}

# ============================================================================
# DATA: DROPDOWN OPTIONS
# ============================================================================

# TODO: Verify these match Excel cell F20 exactly
HVAC_SYSTEMS = [
    'Packaged VAV with electric reheat',
    'Packaged VAV with gas heat',
    'Packaged VAV with PFP boxes',
    'Built-up VAV with hydronic reheat',
    'Built-up VAV with electric reheat',
    'Packaged single zone - AC',
    'Packaged single zone - HP'
]

HEATING_FUELS = ['Electric', 'Natural Gas', 'None']

COOLING_OPTIONS = ['Yes', 'No']

WINDOW_TYPES = ['Single pane', 'Double pane', 'Double pane, low-e']

CSW_TYPES = ['Double', 'Triple', 'Quad']

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
        return "Double"
    return "Double"

def interpolate_operating_hours(operating_hours, val_8760, val_2080):
    """Interpolate between 2080 and 8760 hour values"""
    if operating_hours < 2080:
        return val_2080
    elif operating_hours >= 8760:
        return val_8760
    else:
        return ((operating_hours - 2080) / (8760 - 2080)) * (val_8760 - val_2080) + val_2080

def calculate_savings(inputs):
    """Main calculation function - implements Excel formulas"""
    
    building_area = inputs['building_area']
    csw_area = inputs['csw_area']
    operating_hours = inputs['operating_hours']
    heating_fuel = inputs['heating_fuel']
    electric_rate = inputs['electric_rate']
    gas_rate = inputs['gas_rate']
    state = inputs['state']
    city = inputs['city']
    
    # Get weather data
    weather = WEATHER_DATA_BY_STATE.get(state, {}).get(city, {'HDD': 0, 'CDD': 0})
    hdd = weather['HDD']
    cdd = weather['CDD']
    
    # Simplified calculations - using placeholders
    # TODO: Implement exact Excel formulas with savings lookup table
    
    heating_savings_per_sf = 3.16  # Placeholder
    cooling_savings_per_sf = 3.99  # Placeholder
    
    if inputs['cooling_installed'] == "No":
        cooling_savings_per_sf = 0
    
    total_electric_per_sf = heating_savings_per_sf + cooling_savings_per_sf
    electric_savings_kwh = total_electric_per_sf * csw_area
    
    if heating_fuel == "Natural Gas":
        gas_savings_therms = 0.59 * csw_area
    else:
        gas_savings_therms = 0
    
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    
    baseline_eui = 140.38  # Placeholder
    
    # Calculate % EUI Savings (Excel F54 = Z26)
    # This is the most important metric
    percent_eui_savings = (total_savings_kbtu_sf / baseline_eui) * 100 if baseline_eui > 0 else 0
    
    # WWR calculation - only if CSW area is input
    if csw_area > 0:
        wwr = csw_area / building_area  # Simplified - needs actual regression
    else:
        wwr = None
    
    return {
        'electric_savings_kwh': electric_savings_kwh,
        'gas_savings_therms': gas_savings_therms,
        'electric_cost_savings': electric_cost_savings,
        'gas_cost_savings': gas_cost_savings,
        'total_cost_savings': total_cost_savings,
        'total_savings_kbtu_sf': total_savings_kbtu_sf,
        'baseline_eui': baseline_eui,
        'percent_eui_savings': percent_eui_savings,
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
        state = st.selectbox(
            'Select State',
            options=sorted(WEATHER_DATA_BY_STATE.keys()),
            key='state'
        )
        
        if state:
            city = st.selectbox(
                'Select City',
                options=sorted(WEATHER_DATA_BY_STATE[state].keys()),
                key='city'
            )
    
    if state and city:
        with col2:
            weather = WEATHER_DATA_BY_STATE[state][city]
            st.info(f"**Location HDD (Base 65):** {weather['HDD']:,}")
            st.info(f"**Location CDD (Base 65):** {weather['CDD']:,}")
    
    if st.button('Next →', type='primary'):
        if state and city:
            st.session_state.step = 2
            st.rerun()
        else:
            st.error("Please select both state and city")

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
        
        # Display WWR only after CSW area is entered
        if csw_area > 0 and st.session_state.get('building_area'):
            building_area = st.session_state.get('building_area')
            wwr = csw_area / building_area
            st.info(f"**Est. Window-to-Wall Ratio:** {wwr:.1%}")
    
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
        'state': st.session_state.get('state'),
        'city': st.session_state.get('city'),
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
    
    # ========================================================================
    # PRIMARY RESULT: % EUI SAVINGS (Most Important - Excel F54)
    # ========================================================================
    
    st.markdown('### 🎯 Primary Result')
    
    # Large, prominent display for the most important metric
    col_main = st.columns([1])[0]
    with col_main:
        st.metric(
            label='Energy Use Intensity (EUI) Savings',
            value=f"{results['percent_eui_savings']:.1f}%",
            delta="Energy Reduction",
            help="Percentage reduction in energy use intensity (Excel cell F54)"
        )
    
    st.markdown('---')
    
    # Display intermediary calculations
    with st.expander('📊 Project Details', expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            if results['wwr'] is not None:
                st.metric('Est. Window-to-Wall Ratio', f"{results['wwr']:.1%}")
            else:
                st.metric('Est. Window-to-Wall Ratio', 'N/A')
        with col2:
            st.metric('Location HDD (Base 65)', f"{results['hdd']:,}")
        with col3:
            st.metric('Location CDD (Base 65)', f"{results['cdd']:,}")
    
    st.markdown('---')
    
    # Secondary results
    st.markdown('### 💰 Annual Energy & Cost Savings')
    
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
            '% EUI Savings (PRIMARY)',
            'Electric Energy Savings',
            'Gas Energy Savings',
            'Electric Cost Savings',
            'Gas Cost Savings',
            'Total Savings (Intensity)',
            'Total Cost Savings',
            'Baseline Energy Use Intensity'
        ],
        'Value': [
            f'{results["percent_eui_savings"]:.1f}%',
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
        state = st.session_state.get('state', 'Not set')
        city = st.session_state.get('city', 'Not set')
        st.markdown(f"**Location:** {city}, {state}")
    
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
    st.markdown('This calculator uses simplified savings calculations. Complete Excel data extraction is in progress for exact accuracy.')
    
    st.markdown('---')
    st.markdown('#### 🔧 Next Steps')
    st.markdown("""
    **To Complete:**
    1. Verify HDD/CDD values match Excel exactly
    2. Verify HVAC system options match Excel F20
    3. Extract complete savings lookup table
    4. Implement exact formulas for % EUI savings
    """)

# ============================================================================
# FOOTER
# ============================================================================

st.markdown('---')
st.markdown(
    '<div style="text-align: center; color: gray; font-size: 0.8em;">'
    'CSW Savings Calculator v2.0 | Office Buildings Only'
    '</div>',
    unsafe_allow_html=True
)
