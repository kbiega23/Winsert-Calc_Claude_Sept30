"""
CSW Savings Calculator - Streamlit Web App
REGRESSION-BASED VERSION - Calculates dynamically for each city
"""

import streamlit as st
import pandas as pd
import os
import plotly.graph_objects as go

# ============================================================================
# CONFIGURATION
# ============================================================================

st.set_page_config(
    page_title="Winsert Savings Calculator",
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
WINDOW_TYPES = ['Single pane', 'Double pane']
CSW_TYPES = ['Winsert Lite', 'Winsert Plus']
CSW_TYPE_MAPPING = {'Winsert Lite': 'Single', 'Winsert Plus': 'Double'}

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
def load_regression_coefficients():
    """Load regression coefficients from CSV"""
    try:
        df = pd.read_csv('regression_coefficients.csv')
        return df
    except FileNotFoundError:
        st.error("⚠️ Regression coefficients file not found")
        return pd.DataFrame()

# Load data
WEATHER_DATA_BY_STATE = load_weather_data()
REGRESSION_COEFFICIENTS = load_regression_coefficients()

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

def build_lookup_config(inputs, hours):
    """Build configuration for finding regression row"""
    if inputs['existing_window'] == 'Single pane':
        base = 'Single'
    else:
        base = 'Double'
    
    csw_product = inputs['csw_type']
    csw_type = CSW_TYPE_MAPPING.get(csw_product, csw_product)
    
    if inputs['building_area'] > 30000 and inputs['hvac_system'] == 'Built-up VAV with hydronic reheat':
        size = 'Large'
    else:
        size = 'Mid'
    
    heating_fuel = inputs['heating_fuel']
    if size == 'Mid':
        if heating_fuel in ['Electric', 'None']:
            hvac_fuel = 'PVAV_Elec'
        else:
            hvac_fuel = 'PVAV_Gas'
    else:
        hvac_fuel = 'VAV'
    
    fuel = 'Electric' if heating_fuel == 'None' else heating_fuel
    
    return {
        'base': base,
        'csw': csw_type,
        'size': size,
        'hvac_fuel': hvac_fuel,
        'fuel': fuel,
        'hours': hours
    }

def find_regression_row(config):
    """Find matching regression row based on configuration"""
    if REGRESSION_COEFFICIENTS.empty:
        return None
    
    mask = (
        (REGRESSION_COEFFICIENTS['base'] == config['base']) &
        (REGRESSION_COEFFICIENTS['csw'] == config['csw']) &
        (REGRESSION_COEFFICIENTS['size'] == config['size']) &
        (REGRESSION_COEFFICIENTS['hvac_fuel'] == config['hvac_fuel']) &
        (REGRESSION_COEFFICIENTS['hours'] == config['hours'])
    )
    
    if pd.notna(config['fuel']):
        mask = mask & (REGRESSION_COEFFICIENTS['fuel'] == config['fuel'])
    
    result = REGRESSION_COEFFICIENTS[mask]
    
    if result.empty:
        return None
    
    return result.iloc[0]

def find_baseline_eui_row(config):
    """Find baseline EUI regression row"""
    if REGRESSION_COEFFICIENTS.empty:
        return None
    
    if config['fuel'] == 'Natural Gas':
        fuel_type = 'Gas'
    else:
        fuel_type = 'Electric'
    
    mask = (
        (REGRESSION_COEFFICIENTS['base'] == config['base']) &
        (REGRESSION_COEFFICIENTS['csw'] == 'N/A') &
        (REGRESSION_COEFFICIENTS['size'] == config['size']) &
        (REGRESSION_COEFFICIENTS['hvac_fuel'] == fuel_type) &
        (REGRESSION_COEFFICIENTS['hours'] == config['hours'])
    )
    
    result = REGRESSION_COEFFICIENTS[mask]
    
    if result.empty:
        mask = (
            (REGRESSION_COEFFICIENTS['base'] == config['base']) &
            (REGRESSION_COEFFICIENTS['csw'].isna() | (REGRESSION_COEFFICIENTS['csw'] == 'N/A')) &
            (REGRESSION_COEFFICIENTS['size'] == config['size']) &
            (REGRESSION_COEFFICIENTS['hvac_fuel'] == fuel_type) &
            (REGRESSION_COEFFICIENTS['fuel'].isna() | (REGRESSION_COEFFICIENTS['fuel'] == 'N/A')) &
            (REGRESSION_COEFFICIENTS['hours'] == config['hours'])
        )
        result = REGRESSION_COEFFICIENTS[mask]
    
    if result.empty:
        return None
    
    return result.iloc[0]

def calculate_from_regression(row, degree_days, is_heating=True):
    """Calculate value using regression formula: value = a + b*DD + c*DD²"""
    if row is None:
        return 0
    
    if is_heating:
        a = row['heat_a']
        b = row['heat_b']
        c = row['heat_c']
    else:
        a = row['cool_a']
        b = row['cool_b']
        c = row['cool_c']
    
    value = a + b * degree_days + c * (degree_days ** 2)
    return value

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
    """Main calculation function using regression formulas"""
    building_area = inputs['building_area']
    csw_area = inputs['csw_area']
    operating_hours = inputs['operating_hours']
    num_floors = inputs['num_floors']
    electric_rate = inputs['electric_rate']
    gas_rate = inputs['gas_rate']
    cooling_installed = inputs['cooling_installed']
    heating_fuel = inputs['heating_fuel']
    
    hdd = inputs.get('hdd', 0)
    cdd = inputs.get('cdd', 0)
    
    q24, q25 = calculate_q24_q25(operating_hours)
    
    config_high = build_lookup_config(inputs, q24)
    config_low = build_lookup_config(inputs, q25)
    
    row_high = find_regression_row(config_high)
    row_low = find_regression_row(config_low)
    
    if row_high is None or row_low is None:
        st.error(f"⚠️ Could not find regression coefficients for configuration")
        return None
    
    if heating_fuel == 'Natural Gas':
        heating_high = calculate_from_regression(row_high, hdd, is_heating=True)
        heating_low = calculate_from_regression(row_low, hdd, is_heating=True)
        gas_savings_high = heating_high
        gas_savings_low = heating_low
        electric_heating_high = 0
        electric_heating_low = 0
    else:
        electric_heating_high = calculate_from_regression(row_high, hdd, is_heating=True)
        electric_heating_low = calculate_from_regression(row_low, hdd, is_heating=True)
        gas_savings_high = 0
        gas_savings_low = 0
    
    cooling_high = calculate_from_regression(row_high, cdd, is_heating=False)
    cooling_low = calculate_from_regression(row_low, cdd, is_heating=False)
    
    if heating_fuel == 'Natural Gas':
        c31 = 0
        c33 = interpolate_hours(operating_hours, gas_savings_high, gas_savings_low, q24, q25)
    else:
        c31 = interpolate_hours(operating_hours, electric_heating_high, electric_heating_low, q24, q25)
        c33 = 0
    
   c32_base = interpolate_hours(operating_hours, cooling_high, cooling_low, q24, q25)
    
    # Apply cooling multiplier based on whether cooling is installed
    if cooling_installed == "Yes":
        w24 = 1.0
    else:
        # When cooling is "No", use the cooling multiplier from regression data
        # This represents savings from fans/ventilation even without dedicated cooling
        w24 = row_high.get('cool_mult_no_cooling', 0.6644)
    
    c32 = c32_base * w24
    
    baseline_row_high = find_baseline_eui_row(config_high)
    baseline_row_low = find_baseline_eui_row(config_low)
    
    if baseline_row_high is None or baseline_row_low is None:
        st.error("⚠️ Could not find baseline EUI coefficients")
        return None
    
    baseline_eui_high = calculate_from_regression(baseline_row_high, hdd, is_heating=True)
    baseline_eui_low = calculate_from_regression(baseline_row_low, hdd, is_heating=True)
    baseline_eui = interpolate_hours(operating_hours, baseline_eui_high, baseline_eui_low, q24, q25)
    
    electric_savings_kwh = (c31 + c32) * csw_area
    gas_savings_therms = c33 * csw_area
    
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    
    new_eui = baseline_eui - total_savings_kbtu_sf
    percent_eui_savings = (total_savings_kbtu_sf / baseline_eui * 100) if baseline_eui > 0 else 0
    
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

# Header with logo and title - cleaner, more compact
col_logo, col_title = st.columns([1, 6])
with col_logo:
    if os.path.exists('logo.png'):
        st.image('logo.png', width=180)
    else:
        st.write("")
with col_title:
    st.markdown("<h1 style='margin-bottom: 0;'>Winsert Savings Calculator</h1>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 1.2em; color: #666; margin-top: 0;'>Office Buildings</p>", unsafe_allow_html=True)

st.markdown('---')

# Check if data loaded
if not WEATHER_DATA_BY_STATE:
    st.error("⚠️ Unable to load weather data.")
    st.stop()

if REGRESSION_COEFFICIENTS.empty:
    st.error("⚠️ Unable to load regression coefficients.")
    st.stop()

progress = st.session_state.step / 4
st.progress(progress)
st.write(f'Step {st.session_state.step} of 4')

if st.session_state.step == 1:
    st.header('Step 1: Project Location')
    
    state_options = sorted(WEATHER_DATA_BY_STATE.keys())
    default_state_idx = 0
    if 'state' in st.session_state and st.session_state.state in state_options:
        default_state_idx = state_options.index(st.session_state.state)
    
    state = st.selectbox('Select State', options=state_options, index=default_state_idx, key='state_select')
    st.session_state.state = state
    
    if state:
        city_options = sorted(WEATHER_DATA_BY_STATE[state].keys())
        default_city_idx = 0
        if 'city' in st.session_state and st.session_state.city in city_options:
            default_city_idx = city_options.index(st.session_state.city)
        
        city = st.selectbox('Select City', options=city_options, index=default_city_idx, key='city_select')
        st.session_state.city = city
        
        if city:
            weather = WEATHER_DATA_BY_STATE[state][city]
            st.session_state.hdd = weather['HDD']
            st.session_state.cdd = weather['CDD']
            
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(
                    f"""<div style='padding: 12px; background-color: #f0f2f6; border-radius: 8px; text-align: center;'>
                    <p style='margin: 0; font-size: 0.9em; color: #666;'>Heating Degree Days</p>
                    <p style='margin: 5px 0 0 0; font-size: 1.4em; font-weight: bold; color: #333;'>{weather['HDD']:,.0f}</p>
                    </div>""",
                    unsafe_allow_html=True
                )
            with col2:
                st.markdown(
                    f"""<div style='padding: 12px; background-color: #f0f2f6; border-radius: 8px; text-align: center;'>
                    <p style='margin: 0; font-size: 0.9em; color: #666;'>Cooling Degree Days</p>
                    <p style='margin: 5px 0 0 0; font-size: 1.4em; font-weight: bold; color: #333;'>{weather['CDD']:,.0f}</p>
                    </div>""",
                    unsafe_allow_html=True
                )
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button('Next →', type='primary'):
        if state and city:
            st.session_state.step = 2
            st.rerun()

elif st.session_state.step == 2:
    st.header('Step 2: Building Envelope Information')
    col1, col2 = st.columns(2)
    
    with col1:
        building_area = st.number_input('Building Area (Sq.Ft.)', min_value=15000, max_value=500000, value=st.session_state.get('building_area', 75000), step=1000, key='building_area_input', help="Must be between 15,000 and 500,000 SF")
        st.session_state.building_area = building_area
        
        num_floors = st.number_input('Number of Floors', min_value=1, max_value=100, value=st.session_state.get('num_floors', 5), key='num_floors_input', help="Must be between 1 and 100 floors")
        st.session_state.num_floors = num_floors
        
        window_types_list = WINDOW_TYPES
        existing_window_idx = 0
        if 'existing_window' in st.session_state and st.session_state.existing_window in window_types_list:
            existing_window_idx = window_types_list.index(st.session_state.existing_window)
        existing_window = st.selectbox('Type of Existing Window', options=window_types_list, index=existing_window_idx, key='existing_window_select')
        st.session_state.existing_window = existing_window
    
    with col2:
        if existing_window == 'Double pane':
            csw_types_list = ['Winsert Lite']
            csw_type_idx = 0
        else:
            csw_types_list = CSW_TYPES
            csw_type_idx = 0
            if 'csw_type' in st.session_state and st.session_state.csw_type in csw_types_list:
                csw_type_idx = csw_types_list.index(st.session_state.csw_type)
        
        csw_type = st.selectbox('Secondary Window Product', options=csw_types_list, index=csw_type_idx, key='csw_type_select')
        st.session_state.csw_type = csw_type
        
        csw_area = st.number_input('Total Sq. Ft of Secondary Windows Installed', min_value=0, max_value=int(building_area * 0.5), value=min(st.session_state.get('csw_area', 12000), int(building_area * 0.5)), step=100, key='csw_area_input', help='Check window to wall ratio. Typical values are between 10%-50%.')
        st.session_state.csw_area = csw_area
        
        if csw_area > 0 and building_area > 0 and num_floors > 0:
            wwr = calculate_wwr(csw_area, building_area, num_floors)
            st.markdown(
                f"""<div style='padding: 12px; background-color: #f0f2f6; border-radius: 8px; margin-top: 10px;'>
                <p style='margin: 0; font-size: 0.9em; color: #666;'>Window-to-Wall Ratio</p>
                <p style='margin: 5px 0 0 0; font-size: 1.4em; font-weight: bold; color: #333;'>{wwr:.0%}</p>
                </div>""",
                unsafe_allow_html=True
            )
            
            # WWR validation warnings
            if wwr > 1.0:
                st.error("⚠️ WWR is larger than physically possible. Please update.")
                st.session_state.wwr_error = True
            elif wwr < 0.10 or wwr > 0.50:
                st.warning("⚠️ Warning: window to wall ratio seems out of norm. Please confirm before proceeding.")
                st.session_state.wwr_error = False
            else:
                st.session_state.wwr_error = False
        else:
            st.session_state.wwr_error = False
    
    st.markdown("<br>", unsafe_allow_html=True)
    col_back, col_next = st.columns([1, 1])
    with col_back:
        if st.button('← Back'):
            st.session_state.step = 1
            st.rerun()
    with col_next:
        # Disable next button if WWR > 100%
        can_proceed = not st.session_state.get('wwr_error', False)
        if st.button('Next →', type='primary', disabled=not can_proceed):
            st.session_state.step = 3
            st.rerun()

elif st.session_state.step == 3:
    st.header('Step 3: HVAC & Utility Information')
    col1, col2 = st.columns(2)
    
    with col1:
        electric_rate = st.number_input('Electric Rate ($/kWh)', min_value=0.01, max_value=1.0, value=st.session_state.get('electric_rate', 0.12), step=0.01, format='%.3f', key='electric_rate_input')
        st.session_state.electric_rate = electric_rate
        
        gas_rate = st.number_input('Natural Gas Rate ($/therm)', min_value=0.01, max_value=10.0, value=st.session_state.get('gas_rate', 0.80), step=0.05, format='%.2f', key='gas_rate_input')
        st.session_state.gas_rate = gas_rate
        
        operating_hours = st.number_input('Annual Operating Hours', min_value=1980, max_value=8760, value=st.session_state.get('operating_hours', 8000), step=100, key='operating_hours_input', help="Must be between 1,980 and 8,760 hours")
        st.session_state.operating_hours = operating_hours
    
    with col2:
        hvac_systems_list = HVAC_SYSTEMS
        hvac_idx = 0
        if 'hvac_system' in st.session_state and st.session_state.hvac_system in hvac_systems_list:
            hvac_idx = hvac_systems_list.index(st.session_state.hvac_system)
        hvac_system = st.selectbox('HVAC System Type', options=hvac_systems_list, index=hvac_idx, key='hvac_system_select')
        st.session_state.hvac_system = hvac_system
        
        if hvac_system == 'Packaged VAV with electric reheat':
            heating_fuels_list = ['Electric']
            fuel_idx = 0
        else:
            heating_fuels_list = HEATING_FUELS
            fuel_idx = 0
            if 'heating_fuel' in st.session_state and st.session_state.heating_fuel in heating_fuels_list:
                fuel_idx = heating_fuels_list.index(st.session_state.heating_fuel)
        
        heating_fuel = st.selectbox('Heating Fuel', options=heating_fuels_list, index=fuel_idx, key='heating_fuel_select')
        st.session_state.heating_fuel = heating_fuel
        
        cooling_options_list = COOLING_OPTIONS
        cooling_idx = 0
        if 'cooling_installed' in st.session_state and st.session_state.cooling_installed in cooling_options_list:
            cooling_idx = cooling_options_list.index(st.session_state.cooling_installed)
        cooling_installed = st.selectbox('Cooling Installed?', options=cooling_options_list, index=cooling_idx, key='cooling_installed_select')
        st.session_state.cooling_installed = cooling_installed
    
    st.markdown("<br>", unsafe_allow_html=True)
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
        'hdd': st.session_state.get('hdd', 0),
        'cdd': st.session_state.get('cdd', 0),
        'building_area': st.session_state.get('building_area', 75000),
        'num_floors': st.session_state.get('num_floors', 5),
        'operating_hours': st.session_state.get('operating_hours', 8000),
        'hvac_system': st.session_state.get('hvac_system', HVAC_SYSTEMS[0]),
        'heating_fuel': st.session_state.get('heating_fuel', 'Electric'),
        'cooling_installed': st.session_state.get('cooling_installed', 'Yes'),
        'existing_window': st.session_state.get('existing_window', 'Single pane'),
        'csw_type': st.session_state.get('csw_type', 'Winsert Lite'),
        'csw_area': st.session_state.get('csw_area', 12000),
        'electric_rate': st.session_state.get('electric_rate', 0.12),
        'gas_rate': st.session_state.get('gas_rate', 0.80)
    }
    
    results = calculate_savings(inputs)
    
    if results:
        st.success('✅ Calculation Complete!')
        
        col_chart, col_cost = st.columns([1.3, 1])
        
        with col_chart:
            st.markdown('#### Energy Use Intensity (EUI) Reduction')
            
            baseline_eui = results['baseline_eui']
            savings_eui = results['total_savings_kbtu_sf']
            new_eui = results['new_eui']
            
            fig = go.Figure(go.Waterfall(
                orientation = "v",
                measure = ["absolute", "relative", "total"],
                x = ["Baseline EUI<br>Before Winsert", "Savings with<br>Winsert", "EUI After<br>Winsert"],
                y = [baseline_eui, -savings_eui, new_eui],
                text = [f"{baseline_eui:.1f}", f"−{savings_eui:.1f}", f"{new_eui:.1f}"],
                textposition = "outside",
                textfont = dict(size=14, family='Arial Black'),
                decreasing = {"marker":{"color":"#38ef7d"}},
                increasing = {"marker":{"color":"#FF6B6B"}},
                totals = {"marker":{"color":"#4ECDC4"}},
                connector = {"line":{"color":"rgb(63, 63, 63)"}},
            ))
            
            fig.update_layout(
                height=350,
                showlegend=False,
                yaxis=dict(
                    title='kBtu/SF-yr',
                    title_font=dict(size=12),
                    gridcolor='#E0E0E0'
                ),
                xaxis=dict(
                    title_font=dict(size=12)
                ),
                plot_bgcolor='white',
                paper_bgcolor='white',
                margin=dict(t=20, b=40, l=60, r=20)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown(
                f"""
                <div style='background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); 
                            padding: 20px; 
                            border-radius: 10px; 
                            text-align: center;
                            box-shadow: 0 3px 5px rgba(0,0,0,0.1);'>
                    <h2 style='color: white; margin: 0; font-size: 2.2em; font-weight: bold;'>
                        {results['percent_eui_savings']:.1f}%
                    </h2>
                    <p style='color: white; margin: 5px 0 0 0; font-size: 0.95em;'>
                        EUI Reduction ({results['total_savings_kbtu_sf']:.1f} kBtu/SF-yr)
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )
        
        with col_cost:
            st.markdown('#### Annual Cost Savings')
            
            st.markdown(
                f"""
                <div style='background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                            padding: 25px; 
                            border-radius: 10px; 
                            text-align: center;
                            box-shadow: 0 3px 5px rgba(0,0,0,0.1);
                            margin-bottom: 15px;'>
                    <h1 style='color: white; margin: 0; font-size: 2.5em; font-weight: bold;'>
                        ${results['total_cost_savings']:,.0f}
                    </h1>
                    <p style='color: #E0E0E0; margin: 8px 0 0 0; font-size: 1em;'>
                        Total Annual Savings
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            st.markdown(
                f"""
                <div style='background: #FFF3E0; 
                            padding: 15px; 
                            border-radius: 8px; 
                            margin-bottom: 12px;
                            border-left: 4px solid #FF9800;'>
                    <p style='margin: 0 0 5px 0; color: #E65100; font-size: 0.9em; font-weight: 600;'>⚡ Electric Savings</p>
                    <p style='font-size: 1.6em; margin: 0; font-weight: bold; color: #E65100;'>
                        ${results['electric_cost_savings']:,.0f}<span style='font-size: 0.5em;'>/year</span>
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )
            
            st.markdown(
                f"""
                <div style='background: #E3F2FD; 
                            padding: 15px; 
                            border-radius: 8px;
                            border-left: 4px solid #2196F3;'>
                    <p style='margin: 0 0 5px 0; color: #0D47A1; font-size: 0.9em; font-weight: 600;'>🔥 Natural Gas Savings</p>
                    <p style='font-size: 1.6em; margin: 0; font-weight: bold; color: #0D47A1;'>
                        ${results['gas_cost_savings']:,.0f}<span style='font-size: 0.5em;'>/year</span>
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )
        
        st.markdown('---')
        st.markdown('#### Energy Savings Breakdown')
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(
                f"""<div style='padding: 15px; background-color: #f0f2f6; border-radius: 8px;'>
                <p style='margin: 0; font-size: 0.9em; color: #666;'>Electric Energy Savings</p>
                <p style='margin: 5px 0 0 0; font-size: 1.5em; font-weight: bold; color: #333;'>{results["electric_savings_kwh"]:,.0f} <span style='font-size: 0.6em;'>kWh/yr</span></p>
                </div>""",
                unsafe_allow_html=True
            )
        with col2:
            st.markdown(
                f"""<div style='padding: 15px; background-color: #f0f2f6; border-radius: 8px;'>
                <p style='margin: 0; font-size: 0.9em; color: #666;'>Natural Gas Savings</p>
                <p style='margin: 5px 0 0 0; font-size: 1.5em; font-weight: bold; color: #333;'>{results["gas_savings_therms"]:,.0f} <span style='font-size: 0.6em;'>therms/yr</span></p>
                </div>""",
                unsafe_allow_html=True
            )
        
        st.markdown("<br>", unsafe_allow_html=True)
        with st.expander('🔍 View Detailed Calculations'):
            st.markdown('**Performance Metrics**')
            detail_col1, detail_col2 = st.columns(2)
            with detail_col1:
                st.write(f"• Baseline EUI: {results['baseline_eui']:.1f} kBtu/SF-yr")
                st.write(f"• New EUI: {results['new_eui']:.1f} kBtu/SF-yr")
                st.write(f"• EUI Reduction: {results['percent_eui_savings']:.1f}%")
            with detail_col2:
                st.write(f"• Electric Heating: {results['heating_per_sf']:.4f} kWh/SF-CSW")
                st.write(f"• Cooling & Fans: {results['cooling_per_sf']:.4f} kWh/SF-CSW")
                st.write(f"• Gas Heating: {results['gas_per_sf']:.4f} therms/SF-CSW")
            
            st.markdown('**Project Details**')
            detail_col3, detail_col4 = st.columns(2)
            with detail_col3:
                st.write(f"• Location: {inputs['city']}, {inputs['state']}")
                st.write(f"• Heating Degree Days: {results['hdd']:,.0f}")
                st.write(f"• Cooling Degree Days: {results['cdd']:,.0f}")
            with detail_col4:
                st.write(f"• Building Area: {inputs['building_area']:,} SF")
                st.write(f"• Secondary Window Area: {inputs['csw_area']:,} SF")
                if results['wwr']:
                    st.write(f"• Window-to-Wall Ratio: {results['wwr']:.0%}")
    
    st.markdown("<br>", unsafe_allow_html=True)
    col_back, col_space = st.columns([1, 3])
    with col_back:
        if st.button('← Start Over', type='secondary'):
            keys_to_keep = []
            keys_to_delete = [key for key in st.session_state.keys() if key not in keys_to_keep]
            for key in keys_to_delete:
                del st.session_state[key]
            st.session_state.step = 1
            st.rerun()

with st.sidebar:
    if st.session_state.step == 4:
        st.markdown('### 🎛️ Adjust Inputs')
        st.markdown('Modify values below to see updated results:')
        st.markdown('---')
        
        st.markdown(f"**📍 Location**")
        st.text(f"{st.session_state.get('city', 'N/A')}, {st.session_state.get('state', 'N/A')}")
        st.markdown('---')
        
        st.markdown('**🏢 Building Envelope**')
        building_area = st.number_input('Building Area (SF)', min_value=15000, max_value=500000, value=st.session_state.get('building_area', 75000), step=1000, key='sidebar_building_area')
        if building_area != st.session_state.get('building_area'):
            st.session_state.building_area = building_area
            st.rerun()
        
        num_floors = st.number_input('Floors', min_value=1, max_value=100, value=st.session_state.get('num_floors', 5), key='sidebar_num_floors')
        if num_floors != st.session_state.get('num_floors'):
            st.session_state.num_floors = num_floors
            st.rerun()
        
        existing_window = st.selectbox('Existing Window', options=WINDOW_TYPES, index=WINDOW_TYPES.index(st.session_state.get('existing_window', 'Single pane')), key='sidebar_existing_window')
        if existing_window != st.session_state.get('existing_window'):
            st.session_state.existing_window = existing_window
            st.rerun()
        
        csw_type = st.selectbox('Product', options=CSW_TYPES, index=CSW_TYPES.index(st.session_state.get('csw_type', 'Winsert Lite')), key='sidebar_csw_type')
        if csw_type != st.session_state.get('csw_type'):
            st.session_state.csw_type = csw_type
            st.rerun()
        
        csw_area = st.number_input('Secondary Window Area (SF)', min_value=0, max_value=int(building_area * 0.5), value=min(st.session_state.get('csw_area', 12000), int(building_area * 0.5)), step=100, key='sidebar_csw_area')
        if csw_area != st.session_state.get('csw_area'):
            st.session_state.csw_area = csw_area
            st.rerun()
        
        if csw_area > 0 and building_area > 0 and num_floors > 0:
            wwr = calculate_wwr(csw_area, building_area, num_floors)
            st.text(f"WWR: {wwr:.0%}")
        
        st.markdown('---')
        
        st.markdown('**⚙️ HVAC & Utility**')
        operating_hours = st.number_input('Operating Hours/yr', min_value=1980, max_value=8760, value=st.session_state.get('operating_hours', 8000), step=100, key='sidebar_operating_hours')
        if operating_hours != st.session_state.get('operating_hours'):
            st.session_state.operating_hours = operating_hours
            st.rerun()
        
        hvac_system = st.selectbox('HVAC System', options=HVAC_SYSTEMS, index=HVAC_SYSTEMS.index(st.session_state.get('hvac_system', HVAC_SYSTEMS[0])), key='sidebar_hvac_system')
        if hvac_system != st.session_state.get('hvac_system'):
            st.session_state.hvac_system = hvac_system
            st.rerun()
        
        heating_fuel = st.selectbox('Heating Fuel', options=HEATING_FUELS, index=HEATING_FUELS.index(st.session_state.get('heating_fuel', 'Electric')), key='sidebar_heating_fuel')
        if heating_fuel != st.session_state.get('heating_fuel'):
            st.session_state.heating_fuel = heating_fuel
            st.rerun()
        
        cooling_installed = st.selectbox('Cooling?', options=COOLING_OPTIONS, index=COOLING_OPTIONS.index(st.session_state.get('cooling_installed', 'Yes')), key='sidebar_cooling_installed')
        if cooling_installed != st.session_state.get('cooling_installed'):
            st.session_state.cooling_installed = cooling_installed
            st.rerun()
        
        electric_rate = st.number_input('Electric Rate ($/kWh)', min_value=0.01, max_value=1.0, value=st.session_state.get('electric_rate', 0.12), step=0.01, format='%.3f', key='sidebar_electric_rate')
        if electric_rate != st.session_state.get('electric_rate'):
            st.session_state.electric_rate = electric_rate
            st.rerun()
        
        gas_rate = st.number_input('Gas Rate ($/therm)', min_value=0.01, max_value=10.0, value=st.session_state.get('gas_rate', 0.80), step=0.05, format='%.2f', key='sidebar_gas_rate')
        if gas_rate != st.session_state.get('gas_rate'):
            st.session_state.gas_rate = gas_rate
            st.rerun()
    else:
        st.markdown('### 📝 Summary')
        if st.session_state.step > 1:
            st.markdown(f"**Location:** {st.session_state.get('city', 'N/A')}, {st.session_state.get('state', 'N/A')}")
        if st.session_state.step > 2:
            st.markdown(f"**Building:** {st.session_state.get('building_area', 0):,} SF, {st.session_state.get('num_floors', 0)} floors")
            st.markdown(f"**Windows:** {st.session_state.get('existing_window', 'N/A')} → {st.session_state.get('csw_type', 'N/A')}")
            st.markdown(f"**Secondary Window Area:** {st.session_state.get('csw_area', 0):,} SF")
        if st.session_state.step > 3:
            st.markdown(f"**HVAC:** {st.session_state.get('hvac_system', 'N/A')}")
            st.markdown(f"**Heating:** {st.session_state.get('heating_fuel', 'N/A')}")
            st.markdown(f"**Operating Hours:** {st.session_state.get('operating_hours', 0):,}/yr")
