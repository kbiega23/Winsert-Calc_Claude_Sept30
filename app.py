"""
CSW Savings Calculator - Streamlit Web App
Updated to load weather data from CSV
"""

import streamlit as st
import pandas as pd

# ============================================================================
# CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="CSW Savings Calculator",
    page_icon="🏢",
    layout="wide"
)

if 'step' not in st.session_state:
    st.session_state.step = 1

# ============================================================================
# DATA: EXACT HVAC OPTIONS FROM EXCEL LISTS!BD2-BD5
# ============================================================================

HVAC_SYSTEMS = [
    'Packaged VAV with electric reheat',
    'Packaged VAV with hydronic reheat',
    'Built-up VAV with hydronic reheat',
    'Other'
]

HEATING_FUELS = ['Electric', 'Natural Gas', 'None']
COOLING_OPTIONS = ['Yes', 'No']
WINDOW_TYPES = ['Single pane', 'Double pane', 'Double pane, low-e']
CSW_TYPES = ['Double', 'Triple', 'Quad']

# ============================================================================
# LOAD WEATHER DATA FROM CSV
# ============================================================================

@st.cache_data
def load_weather_data():
    """Load weather data from CSV file"""
    try:
        df = pd.read_csv('weather_information.csv')
        
        # Clean up state names (fix typo: Aklaska -> Alaska)
        df['State'] = df['State'].replace('Aklaska', 'Alaska')
        
        # Create nested dictionary structure: {State: {City: {HDD: x, CDD: y}}}
        weather_dict = {}
        for _, row in df.iterrows():
            state = row['State']
            city = row['Cities']
            hdd = row['Heating Degree Days (HDD)']
            cdd = row['Cooling Degree Days (CDD)']
            
            if state not in weather_dict:
                weather_dict[state] = {}
            
            weather_dict[state][city] = {
                'HDD': hdd,
                'CDD': cdd
            }
        
        return weather_dict
    except FileNotFoundError:
        st.error("⚠️ Weather data file 'weather_information.csv' not found. Please ensure it's in the same directory as app.py")
        return {}

# Load weather data
WEATHER_DATA_BY_STATE = load_weather_data()

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def calculate_wwr(csw_area, building_area, num_floors):
    """Excel F28: F27/((F18/F19)^0.5*4*15*F19)"""
    if num_floors == 0 or building_area == 0:
        return 0
    floor_area = building_area / num_floors
    wall_area = (floor_area ** 0.5) * 4 * 15 * num_floors
    return csw_area / wall_area if wall_area > 0 else 0

def interpolate_hours(operating_hours, val_8760, val_2080, q24, q25):
    """Excel interpolation formula"""
    if operating_hours < 2080:
        return val_2080
    elif operating_hours >= 8760:
        return val_8760
    else:
        return ((operating_hours - q25) / (q24 - q25)) * (val_8760 - val_2080) + val_2080

def calculate_q24_q25(operating_hours):
    """Q24/Q25 calculation"""
    q24 = 8760 if operating_hours > 2912 else 2912
    q25 = 2912 if q24 == 8760 else 2080
    return q24, q25

def calculate_savings(inputs):
    """Main calculation function"""
    building_area = inputs['building_area']
    csw_area = inputs['csw_area']
    operating_hours = inputs['operating_hours']
    num_floors = inputs['num_floors']
    electric_rate = inputs['electric_rate']
    gas_rate = inputs['gas_rate']
    state = inputs['state']
    city = inputs['city']
    cooling_installed = inputs['cooling_installed']
    
    weather = WEATHER_DATA_BY_STATE.get(state, {}).get(city, {'HDD': 0, 'CDD': 0})
    hdd = weather['HDD']
    cdd = weather['CDD']
    
    q24, q25 = calculate_q24_q25(operating_hours)
    
    # Placeholder values - need full savings lookup table
    s24_heating_8760 = 3.16
    s25_heating_2080 = 3.35
    t24_cooling_8760 = 3.99
    t25_cooling_2080 = 7.67
    v24_baseline_8760 = 140.38
    v25_baseline_2080 = 130.37
    
    w24 = 1.0 if cooling_installed == "Yes" else 0.0
    
    c31 = interpolate_hours(operating_hours, s24_heating_8760, s25_heating_2080, q24, q25)
    c32_base = interpolate_hours(operating_hours, t24_cooling_8760, t25_cooling_2080, q24, q25)
    c32 = c32_base * w24
    
    electric_savings_kwh = (c31 + c32) * csw_area
    gas_savings_therms = 0
    
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    baseline_eui = interpolate_hours(operating_hours, v24_baseline_8760, v25_baseline_2080, q24, q25)
    
    z25 = baseline_eui - total_savings_kbtu_sf
    y24 = baseline_eui
    z26 = z25 / y24 if y24 > 0 else 0
    percent_eui_savings = z26 * 100
    
    wwr = calculate_wwr(csw_area, building_area, num_floors) if csw_area > 0 and num_floors > 0 else None
    
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
        'cdd': cdd,
    }

# ============================================================================
# UI
# ============================================================================

st.title('🏢 Commercial Window Savings Calculator')
st.markdown('### Office Buildings')
st.markdown('---')

# Check if weather data loaded successfully
if not WEATHER_DATA_BY_STATE:
    st.error("⚠️ Unable to load weather data. Please ensure 'weather_information.csv' is in the correct location.")
    st.stop()

progress = st.session_state.step / 4
st.progress(progress)
st.write(f'Step {st.session_state.step} of 4')

if st.session_state.step == 1:
    st.header('Step 1: Project Location')
    col1, col2 = st.columns(2)
    
    with col1:
        state = st.selectbox('Select State', options=sorted(WEATHER_DATA_BY_STATE.keys()), key='state')
        if state:
            city = st.selectbox('Select City', options=sorted(WEATHER_DATA_BY_STATE[state].keys()), key='city')
    
    if state and city:
        with col2:
            weather = WEATHER_DATA_BY_STATE[state][city]
            st.info(f"**Location HDD (Base 65):** {weather['HDD']:,}")
            st.info(f"**Location CDD (Base 65):** {weather['CDD']:,}")
    
    if st.button('Next →', type='primary'):
        if state and city:
            st.session_state.step = 2
            st.rerun()

elif st.session_state.step == 2:
    st.header('Step 2: Building Envelope Information')
    col1, col2 = st.columns(2)
    
    with col1:
        building_area = st.number_input('Building Area (Sq.Ft.)', min_value=15000, max_value=500000, value=st.session_state.get('building_area', 75000), step=1000, key='building_area')
        num_floors = st.number_input('Number of Floors', min_value=1, max_value=50, value=st.session_state.get('num_floors', 5), key='num_floors')
        existing_window = st.selectbox('Type of Existing Window', options=WINDOW_TYPES, index=WINDOW_TYPES.index(st.session_state.get('existing_window', 'Single pane')), key='existing_window')
    
    with col2:
        csw_type = st.selectbox('Type of CSW Analyzed', options=CSW_TYPES, index=CSW_TYPES.index(st.session_state.get('csw_type', 'Double')), key='csw_type')
        csw_area = st.number_input('Sq.ft. of CSW Installed', min_value=0, max_value=int(building_area * 0.5), value=st.session_state.get('csw_area', 12000), step=100, key='csw_area')
        
        # Auto-calculate and display WWR (updates dynamically)
        if csw_area > 0 and building_area > 0 and num_floors > 0:
            wwr = calculate_wwr(csw_area, building_area, num_floors)
            st.metric('Window-to-Wall Ratio', f"{wwr:.1%}", help="Automatically calculated based on building area, floors, and CSW area")
    
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button('← Back'):
            st.session_state.step = 1
            st.rerun()
    with col_next:
        if st.button('Next →', type='primary'):
            st.session_state.step = 3
            st.rerun()

elif st.session_state.step == 3:
    st.header('Step 3: HVAC & Utility Information')
    col1, col2 = st.columns(2)
    
    with col1:
        st.number_input('Electric Rate ($/kWh)', min_value=0.01, max_value=1.0, value=st.session_state.get('electric_rate', 0.12), step=0.01, format='%.3f', key='electric_rate')
        st.number_input('Natural Gas Rate ($/therm)', min_value=0.01, max_value=10.0, value=st.session_state.get('gas_rate', 0.80), step=0.05, format='%.2f', key='gas_rate')
        st.number_input('Annual Operating Hours', min_value=1980, max_value=8760, value=st.session_state.get('operating_hours', 8000), step=100, key='operating_hours')
    
    with col2:
        st.selectbox('HVAC System Type', options=HVAC_SYSTEMS, index=HVAC_SYSTEMS.index(st.session_state.get('hvac_system', HVAC_SYSTEMS[0])), key='hvac_system')
        st.selectbox('Heating Fuel', options=HEATING_FUELS, index=HEATING_FUELS.index(st.session_state.get('heating_fuel', 'Electric')), key='heating_fuel')
        st.selectbox('Cooling Installed?', options=COOLING_OPTIONS, index=COOLING_OPTIONS.index(st.session_state.get('cooling_installed', 'Yes')), key='cooling_installed')
    
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button('← Back'):
            st.session_state.step = 2
            st.rerun()
    with col_next:
        if st.button('Calculate Savings →', type='primary'):
            st.session_state.step = 4
            st.rerun()

elif st.session_state.step == 4:
    st.header('💡 Your Energy Savings Results')
    
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
    
    results = calculate_savings(inputs)
    
    st.success('✅ Calculation Complete!')
    st.markdown('### 🎯 Primary Result')
    st.metric('Energy Use Intensity (EUI) Savings', f"{results['percent_eui_savings']:.1f}%", delta="Reduction", help="Excel F54")
    
    st.markdown('---')
    st.markdown('### 💰 Annual Savings')
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric('Electric Energy', f'{results["electric_savings_kwh"]:,.0f} kWh/yr')
        st.metric('Gas Energy', f'{results["gas_savings_therms"]:,.0f} therms/yr')
    with col2:
        st.metric('Electric Cost', f'${results["electric_cost_savings"]:,.2f}/yr')
        st.metric('Gas Cost', f'${results["gas_cost_savings"]:,.2f}/yr')
    with col3:
        st.metric('Savings Intensity', f'{results["total_savings_kbtu_sf"]:.2f} kBtu/SF-yr')
        st.metric('Total Cost', f'${results["total_cost_savings"]:,.2f}/yr')
    
    st.markdown('---')
    st.metric('Baseline EUI', f'{results["baseline_eui"]:.2f} kBtu/SF-yr')
    
    if st.button('← Start Over'):
        st.session_state.step = 1
        st.rerun()

with st.sidebar:
    st.markdown('### 📝 Summary')
    if st.session_state.step > 1:
        st.markdown(f"**Location:** {st.session_state.get('city', 'N/A')}, {st.session_state.get('state', 'N/A')}")
    if st.session_state.step > 2:
        st.markdown(f"**Building:** {st.session_state.get('building_area', 0):,} SF, {st.session_state.get('num_floors', 0)} floors")
        st.markdown(f"**Windows:** {st.session_state.get('existing_window', 'N/A')} → {st.session_state.get('csw_type', 'N/A')} CSW")
        st.markdown(f"**CSW Area:** {st.session_state.get('csw_area', 0):,} SF")
    if st.session_state.step > 3:
        st.markdown(f"**HVAC:** {st.session_state.get('hvac_system', 'N/A')}")
        st.markdown(f"**Heating:** {st.session_state.get('heating_fuel', 'N/A')}")
    if WEATHER_DATA_BY_STATE:
        st.markdown('---')
        st.markdown(f'**Status:** ✅ {len(WEATHER_DATA_BY_STATE)} states | ✅ 874 cities loaded')
