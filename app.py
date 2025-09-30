"""
CSW Savings Calculator - Streamlit Web App
COMPLETE VERSION with working calculations from Excel
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
# DATA: EXACT HVAC OPTIONS FROM EXCEL
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
# LOAD DATA FROM CSV FILES
# ============================================================================

@st.cache_data
def load_weather_data():
    """Load weather data from CSV file"""
    try:
        df = pd.read_csv('weather_information.csv')
        df['State'] = df['State'].replace('Aklaska', 'Alaska')
        
        weather_dict = {}
        for _, row in df.iterrows():
            state = row['State']
            city = row['Cities']
            hdd = row['Heating Degree Days (HDD)']
            cdd = row['Cooling Degree Days (CDD)']
            
            if state not in weather_dict:
                weather_dict[state] = {}
            
            weather_dict[state][city] = {'HDD': hdd, 'CDD': cdd}
        
        return weather_dict
    except FileNotFoundError:
        st.error("⚠️ Weather data file not found")
        return {}

@st.cache_data
def load_savings_lookup():
    """Load savings lookup table from CSV"""
    try:
        df = pd.read_csv('savings_lookup.csv')
        return df
    except FileNotFoundError:
        st.error("⚠️ Savings lookup file not found")
        return pd.DataFrame()

# Load data
WEATHER_DATA_BY_STATE = load_weather_data()
SAVINGS_LOOKUP = load_savings_lookup()

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

def build_lookup_key(inputs, hours):
    """
    Build the lookup key string matching Excel's concatenation logic
    Format: BaseWindow + CSWType + Size + BuildingType + HVACType + FuelType + Hours
    """
    # Base window type (K24)
    if inputs['existing_window'] == 'Single pane':
        base = 'Single'
    else:
        base = 'Double'
    
    # CSW type (L24)
    csw_type = inputs['csw_type']
    
    # Size (M24)
    if inputs['building_area'] > 30000 and inputs['hvac_system'] == 'Built-up VAV with hydronic reheat':
        size = 'Large'
    else:
        size = 'Mid'
    
    # Building type (N24)
    building_type = 'Office'
    
    # HVAC System Type (O24)
    heating_fuel = inputs['heating_fuel']
    if size == 'Mid':
        if heating_fuel in ['Electric', 'None']:
            hvac_type = 'PVAV_Elec'
        else:
            hvac_type = 'PVAV_Gas'
    else:
        hvac_type = 'VAV'
    
    # Fuel type (P24)
    fuel = 'Electric' if heating_fuel == 'None' else heating_fuel
    
    # Combine all parts
    lookup_key = f"{base}{csw_type}{size}{building_type}{hvac_type}{fuel}{hours}"
    
    return lookup_key

def get_savings_from_lookup(lookup_key):
    """Query the savings lookup table for a specific key"""
    if SAVINGS_LOOKUP.empty:
        return None
    
    result = SAVINGS_LOOKUP[SAVINGS_LOOKUP['lookup_key'] == lookup_key]
    
    if result.empty:
        return None
    
    return {
        'heating_kwh': result.iloc[0]['heating_kwh'],
        'cooling_kwh': result.iloc[0]['cooling_kwh'],
        'gas_therms': result.iloc[0]['gas_therms'],
        'baseline_eui': result.iloc[0]['baseline_eui']
    }

def interpolate_hours(operating_hours, val_8760, val_2080_or_2912, q24, q25):
    """Excel interpolation formula"""
    if operating_hours < 2080:
        return val_2080_or_2912
    elif operating_hours >= 8760:
        return val_8760
    else:
        return ((operating_hours - q25) / (q24 - q25)) * (val_8760 - val_2080_or_2912) + val_2080_or_2912

def calculate_q24_q25(operating_hours):
    """Q24/Q25 calculation"""
    if operating_hours > 2912:
        q24 = 8760
        q25 = 2912
    else:
        q24 = 2912
        q25 = 2080
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
    
    # Get weather data
    weather = WEATHER_DATA_BY_STATE.get(state, {}).get(city, {'HDD': 0, 'CDD': 0})
    hdd = weather['HDD']
    cdd = weather['CDD']
    
    # Determine interpolation points
    q24, q25 = calculate_q24_q25(operating_hours)
    
    # Build lookup keys
    key_high = build_lookup_key(inputs, q24)
    key_low = build_lookup_key(inputs, q25)
    
    # Get savings values
    values_high = get_savings_from_lookup(key_high)
    values_low = get_savings_from_lookup(key_low)
    
    if not values_high or not values_low:
        st.error(f"⚠️ Lookup keys not found: {key_high} or {key_low}")
        return None
    
    # Interpolate heating savings
    s_high = values_high['heating_kwh']
    s_low = values_low['heating_kwh']
    c31 = interpolate_hours(operating_hours, s_high, s_low, q24, q25)
    
    # Interpolate cooling savings
    t_high = values_high['cooling_kwh']
    t_low = values_low['cooling_kwh']
    c32_base = interpolate_hours(operating_hours, t_high, t_low, q24, q25)
    
    # Apply cooling multiplier
    w24 = 1.0 if cooling_installed == "Yes" else 0.0
    c32 = c32_base * w24
    
    # Interpolate gas savings
    u_high = values_high['gas_therms']
    u_low = values_low['gas_therms']
    c33 = interpolate_hours(operating_hours, u_high, u_low, q24, q25)
    
    # Interpolate baseline EUI
    v_high = values_high['baseline_eui']
    v_low = values_low['baseline_eui']
    baseline_eui = interpolate_hours(operating_hours, v_high, v_low, q24, q25)
    
    # Calculate total energy savings
    electric_savings_kwh = (c31 + c32) * csw_area
    gas_savings_therms = c33 * csw_area
    
    # Calculate cost savings
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    # Calculate savings intensity
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    
    # Calculate EUI savings percentage
    new_eui = baseline_eui - total_savings_kbtu_sf
    percent_eui_savings = (total_savings_kbtu_sf / baseline_eui * 100) if baseline_eui > 0 else 0
    
    # Calculate WWR
    wwr = calculate_wwr(csw_area, building_area, num_floors) if csw_area > 0 and num_floors > 0 else None
    
    return {
        'electric_savings_kwh': electric_savings_kwh,
        'gas_savings_therms': gas_savings_therms,
        'electric_cost_savings': electric_cost_savings,
        'gas_cost_savings': gas_cost_savings,
        'total_cost_savings': total_cost_savings,
        'total_savings_kbtu_sf': total_savings_kbtu_sf,
        'baseline_eui': baseline_eui,
        'new_eui': new_eui,
        'percent_eui_savings': percent_eui_savings,
        'wwr': wwr,
        'hdd': hdd,
        'cdd': cdd,
        'heating_per_sf': c31,
        'cooling_per_sf': c32,
        'gas_per_sf': c33
    }

# ============================================================================
# UI
# ============================================================================

st.title('🏢 Commercial Window Savings Calculator')
st.markdown('### Office Buildings')
st.markdown('---')

# Check if data loaded
if not WEATHER_DATA_BY_STATE:
    st.error("⚠️ Unable to load weather data.")
    st.stop()

if SAVINGS_LOOKUP.empty:
    st.error("⚠️ Unable to load savings lookup data.")
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
        
        if csw_area > 0 and building_area > 0 and num_floors > 0:
            wwr = calculate_wwr(csw_area, building_area, num_floors)
            st.metric('Window-to-Wall Ratio', f"{wwr:.1%}", help="Automatically calculated")
    
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
    
    if results:
        st.success('✅ Calculation Complete!')
        st.markdown('### 🎯 Primary Result')
        st.metric('Energy Use Intensity (EUI) Savings', f"{results['percent_eui_savings']:.1f}%", delta="Reduction")
        
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
        col1, col2 = st.columns(2)
        with col1:
            st.metric('Baseline EUI', f'{results["baseline_eui"]:.2f} kBtu/SF-yr')
        with col2:
            st.metric('New EUI', f'{results["new_eui"]:.2f} kBtu/SF-yr')
        
        with st.expander('🔍 Calculation Details (Debug Info)'):
            st.write("**Lookup Keys:**")
            q24, q25 = calculate_q24_q25(inputs['operating_hours'])
            key_high = build_lookup_key(inputs, q24)
            key_low = build_lookup_key(inputs, q25)
            st.code(f"High key ({q24} hrs): {key_high}")
            st.code(f"Low key ({q25} hrs): {key_low}")
            
            st.write("**Lookup Values:**")
            values_high = get_savings_from_lookup(key_high)
            values_low = get_savings_from_lookup(key_low)
            if values_high and values_low:
                st.write(f"High values: Heat={values_high['heating_kwh']:.4f}, Cool={values_high['cooling_kwh']:.4f}, Gas={values_high['gas_therms']:.4f}, EUI={values_high['baseline_eui']:.2f}")
                st.write(f"Low values: Heat={values_low['heating_kwh']:.4f}, Cool={values_low['cooling_kwh']:.4f}, Gas={values_low['gas_therms']:.4f}, EUI={values_low['baseline_eui']:.2f}")
            
            st.write("**Interpolated Results:**")
            st.write(f"Heating Savings: {results['heating_per_sf']:.4f} kWh/SF-CSW")
            st.write(f"Cooling Savings: {results['cooling_per_sf']:.4f} kWh/SF-CSW")
            st.write(f"Gas Savings: {results['gas_per_sf']:.4f} therms/SF-CSW")
            st.write(f"Baseline EUI: {results['baseline_eui']:.2f} kBtu/SF-yr")
            
            st.write("**Climate Data:**")
            st.write(f"HDD: {results['hdd']:,}, CDD: {results['cdd']:,}")
            
            if results['wwr']:
                st.write(f"**Window-to-Wall Ratio:** {results['wwr']:.1%}")
    
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
        st.markdown(f"**Operating Hours:** {st.session_state.get('operating_hours', 0):,}/yr")
    if WEATHER_DATA_BY_STATE and not SAVINGS_LOOKUP.empty:
        st.markdown('---')
        st.markdown(f'**Status:** ✅ {len(WEATHER_DATA_BY_STATE)} states | ✅ 874 cities | ✅ {len(SAVINGS_LOOKUP)} lookup combinations')
