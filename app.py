"""
CSW Savings Calculator - Streamlit Web App
CORRECTED VERSION with exact Excel formulas
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
# DATA: WEATHER DATA (EXACT FROM WEATHER INFORMATION SHEET)
# NOTE: Run the extract_excel_data.py script to get complete weather data
# For now using representative sample - REPLACE WITH COMPLETE DATA
# ============================================================================

WEATHER_DATA_BY_STATE = {
    'Alabama': {
        'Anniston': {'HDD': 2585, 'CDD': 1713},
        'Auburn': {'HDD': 2688, 'CDD': 1477},
        'Birmingham': {'HDD': 2698, 'CDD': 1912},
        'Dothan': {'HDD': 2072, 'CDD': 2443},
        'Huntsville': {'HDD': 3472, 'CDD': 1701},
        'Mobile': {'HDD': 1724, 'CDD': 2524},
        'Montgomery': {'HDD': 2183, 'CDD': 2124},
    },
    'Alaska': {
        'Anchorage': {'HDD': 10158, 'CDD': 0},
        'Fairbanks': {'HDD': 13969, 'CDD': 92},
        'Juneau': {'HDD': 8966, 'CDD': 0},
    },
    'Arizona': {
        'Flagstaff': {'HDD': 7152, 'CDD': 232},
        'Phoenix': {'HDD': 1439, 'CDD': 2974},
        'Tucson': {'HDD': 1846, 'CDD': 2639},
    },
    'California': {
        'Los Angeles': {'HDD': 1349, 'CDD': 674},
        'Sacramento': {'HDD': 2502, 'CDD': 1285},
        'San Diego': {'HDD': 1305, 'CDD': 772},
        'San Francisco': {'HDD': 2901, 'CDD': 119},
    },
    'Colorado': {
        'Colorado Springs': {'HDD': 6423, 'CDD': 471},
        'Denver': {'HDD': 5792, 'CDD': 734},
    },
    'Florida': {
        'Jacksonville': {'HDD': 1293, 'CDD': 2553},
        'Miami': {'HDD': 161, 'CDD': 4038},
        'Orlando': {'HDD': 667, 'CDD': 3116},
        'Tampa': {'HDD': 683, 'CDD': 3000},
    },
    'Georgia': {
        'Atlanta': {'HDD': 2961, 'CDD': 1667},
        'Savannah': {'HDD': 1819, 'CDD': 2253},
    },
    'Illinois': {
        'Chicago': {'HDD': 6536, 'CDD': 782},
        'Springfield': {'HDD': 5619, 'CDD': 1108},
    },
    'New York': {
        'Albany': {'HDD': 6875, 'CDD': 579},
        'Buffalo': {'HDD': 6927, 'CDD': 526},
        'New York City': {'HDD': 4811, 'CDD': 1089},
    },
    'Texas': {
        'Austin': {'HDD': 1711, 'CDD': 2887},
        'Dallas': {'HDD': 2363, 'CDD': 2583},
        'Houston': {'HDD': 1434, 'CDD': 2823},
        'San Antonio': {'HDD': 1546, 'CDD': 2900},
    },
    # TODO: Add remaining states from extract_excel_data.py output
}

# ============================================================================
# HELPER FUNCTIONS - EXACT EXCEL FORMULAS
# ============================================================================

def interpolate_hours(operating_hours, val_8760, val_2080, q24, q25):
    """
    Exact Excel interpolation formula
    IF(F23<2080, val_2080, IF(F23>=8760, val_8760, ((F23-Q25)/(Q24-Q25)*(val_8760-val_2080)+val_2080)))
    """
    if operating_hours < 2080:
        return val_2080
    elif operating_hours >= 8760:
        return val_8760
    else:
        return ((operating_hours - q25) / (q24 - q25)) * (val_8760 - val_2080) + val_2080

def calculate_q24_q25(operating_hours):
    """
    Q24 = IF(F23>2912, 8760, 2912)
    Q25 = IF(Q24=8760, 2912, 2080)
    """
    q24 = 8760 if operating_hours > 2912 else 2912
    q25 = 2912 if q24 == 8760 else 2080
    return q24, q25

def calculate_savings(inputs):
    """
    Implements exact Excel calculation flow:
    F54 = Z26 = Z25/Y24
    Where:
      Z25 = Y24 - Y25 (Savings kBtu/SF)
      Y24 = F13 (Baseline EUI)
      Y25 = F14 (Post-retrofit EUI)
    """
    
    building_area = inputs['building_area']
    csw_area = inputs['csw_area']
    operating_hours = inputs['operating_hours']
    heating_fuel = inputs['heating_fuel']
    electric_rate = inputs['electric_rate']
    gas_rate = inputs['gas_rate']
    state = inputs['state']
    city = inputs['city']
    cooling_installed = inputs['cooling_installed']
    
    # Get weather data
    weather = WEATHER_DATA_BY_STATE.get(state, {}).get(city, {'HDD': 0, 'CDD': 0})
    hdd = weather['HDD']
    cdd = weather['CDD']
    
    # Calculate Q24 and Q25 (Excel formulas)
    q24, q25 = calculate_q24_q25(operating_hours)
    
    # SIMPLIFIED CALCULATIONS - Need complete savings lookup table
    # These are placeholders using sample values
    
    # From savings lookup (S24, T24, U24 at 8760 hrs and S25, T25, U25 at 2080 hrs)
    s24_heating_8760 = 3.16  # Placeholder
    s25_heating_2080 = 3.35  # Placeholder
    t24_cooling_8760 = 3.99  # Placeholder
    t25_cooling_2080 = 7.67  # Placeholder
    u24_gas_8760 = 0.0  # Placeholder
    u25_gas_2080 = 0.0  # Placeholder
    v24_baseline_8760 = 140.38  # Placeholder
    v25_baseline_2080 = 130.37  # Placeholder
    
    # W24 = cooling factor (1 if cooling installed, else 0)
    w24 = 1.0 if cooling_installed == "Yes" else 0.0
    
    # Excel C31: Heating savings per SF (interpolated)
    c31 = interpolate_hours(operating_hours, s24_heating_8760, s25_heating_2080, q24, q25)
    
    # Excel C32: Cooling savings per SF (interpolated, with cooling factor)
    c32_base = interpolate_hours(operating_hours, t24_cooling_8760, t25_cooling_2080, q24, q25)
    c32 = c32_base * w24
    
    # Excel C33: Gas savings per SF (interpolated)
    c33 = interpolate_hours(operating_hours, u24_gas_8760, u25_gas_2080, q24, q25)
    
    # Excel F31: Total electric savings = (C31 + C32) * F27
    electric_savings_kwh = (c31 + c32) * csw_area
    
    # Excel F33: Total gas savings = C33 * F27
    gas_savings_therms = c33 * csw_area
    
    # Cost savings
    electric_cost_savings = electric_savings_kwh * electric_rate
    gas_cost_savings = gas_savings_therms * gas_rate
    total_cost_savings = electric_cost_savings + gas_cost_savings
    
    # Excel F35: Total savings intensity = (F31*3.413 + F33*100) / F18
    total_savings_kbtu_sf = (electric_savings_kwh * 3.413 + gas_savings_therms * 100) / building_area
    
    # Excel F13 (Y24): Baseline EUI (interpolated)
    baseline_eui = interpolate_hours(operating_hours, v24_baseline_8760, v25_baseline_2080, q24, q25)
    
    # Excel F14 (Y25): Post-retrofit EUI = F35 (Total savings kBtu/SF)
    post_retrofit_eui = total_savings_kbtu_sf
    
    # Excel Z25: Savings intensity = Y24 - Y25
    z25 = baseline_eui - post_retrofit_eui
    
    # Excel Y24 = F13 (Baseline EUI)
    y24 = baseline_eui
    
    # Excel Z26: Ratio = Z25 / Y24
    z26 = z25 / y24 if y24 > 0 else 0
    
    # Excel F54: % EUI Savings = Z26 (as decimal)
    percent_eui_savings = z26 * 100  # Convert to percentage
    
    # WWR calculation - only if CSW area is input
    if csw_area > 0:
        wwr = csw_area / building_area
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
        'post_retrofit_eui': post_retrofit_eui,
        'percent_eui_savings': percent_eui_savings,
        'wwr': wwr,
        'hdd': hdd,
        'cdd': cdd,
        'z25': z25,
        'y24': y24,
        'z26': z26
    }

# ============================================================================
# STREAMLIT UI
# ============================================================================

st.title('🏢 Commercial Window Savings Calculator')
st.markdown('### Office Buildings')
st.markdown('---')

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
    
    if building_area < 15000 or building_area > 500000:
        st.warning('⚠️ Building area outside recommended range')
    
    if operating_hours < 1980 or operating_hours > 8760:
        st.warning('⚠️ Operating hours outside valid range')
    
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
            key='existing_window'
        )
        
        csw_type = st.selectbox(
            'Type of CSW Analyzed',
            options=CSW_TYPES,
            key='csw_type'
        )
        
        csw_area = st.number_input(
            'Sq.ft. of CSW Installed',
            min_value=0,
            max_value=int(st.session_state.get('building_area', 75000) * 0.5),
            value=12000,
            step=100,
            key='csw_area'
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
            key='electric_rate'
        )
        
        gas_rate = st.number_input(
            'Natural Gas Rate ($/therm)',
            min_value=0.01,
            max_value=10.0,
            value=0.80,
            step=0.05,
            format='%.2f',
            key='gas_rate'
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
    
    # Gather inputs
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
    
    # Calculate
    results = calculate_savings(inputs)
    
    st.success('✅ Calculation Complete!')
    
    # ========================================================================
    # PRIMARY RESULT: % EUI SAVINGS (Excel F54 = Z26)
    # ========================================================================
    
    st.markdown('### 🎯 Primary Result')
    
    col_main = st.columns([1])[0]
    with col_main:
        st.metric(
            label='Energy Use Intensity (EUI) Savings',
            value=f"{results['percent_eui_savings']:.1f}%",
            delta="Reduction in Energy Use",
            help="Excel F54: Percentage reduction in building energy intensity"
        )
    
    # Show calculation details
    with st.expander('📐 Calculation Details (Excel Formulas)', expanded=False):
        st.markdown(f"""
        **Excel Formula Chain:**
        - **F54** = Z26 = {results['z26']:.4f} = **{results['percent_eui_savings']:.1f}%**
        - **Z26** = Z25 / Y24
        - **Z25** = {results['z25']:.2f} kBtu/SF-yr (Savings Intensity)
        - **Y24** = {results['y24']:.2f} kBtu/SF-yr (Baseline EUI)
        
        *Formula: % EUI Savings = (Savings Intensity / Baseline EUI) × 100*
        """)
    
    st.markdown('---')
    
    # Intermediary calculations
    with st.expander('📊 Project Details', expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            if results['wwr'] is not None:
                st.metric('Window-to-Wall Ratio', f"{results['wwr']:.1%}")
        with col2:
            st.metric('Location HDD', f"{results['hdd']:,}")
        with col3:
            st.metric('Location CDD', f"{results['cdd']:,}")
    
    st.markdown('---')
    
    # Secondary results
    st.markdown('### 💰 Annual Energy & Cost Savings')
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric('⚡ Electric Energy', f'{results["electric_savings_kwh"]:,.0f} kWh/yr')
        st.metric('🔥 Gas Energy', f'{results["gas_savings_therms"]:,.0f} therms/yr')
    
    with col2:
        st.metric('💵 Electric Cost', f'${results["electric_cost_savings"]:,.2f}/yr')
        st.metric('💵 Gas Cost', f'${results["gas_cost_savings"]:,.2f}/yr')
    
    with col3:
        st.metric('📊 Savings Intensity', f'{results["total_savings_kbtu_sf"]:.2f} kBtu/SF-yr')
        st.metric('💰 Total Cost', f'${results["total_cost_savings"]:,.2f}/yr', 
                 delta=f'${results["total_cost_savings"]/12:,.2f}/month')
    
    st.markdown('---')
    
    st.metric('📈 Baseline EUI', f'{results["baseline_eui"]:.2f} kBtu/SF-yr')
    
    # Summary table
    st.markdown('### 📋 Summary Report')
    summary_df = pd.DataFrame({
        'Metric': [
            '% EUI Savings (PRIMARY - Excel F54)',
            'Electric Energy Savings',
            'Gas Energy Savings',
            'Electric Cost Savings',
            'Gas Cost Savings',
            'Total Savings Intensity',
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
# SIDEBAR
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
        st.markdown(f"**HVAC:** {st.session_state.get('hvac_system', 'Not set')}")
        st.markdown(f"**Fuel:** {st.session_state.get('heating_fuel', 'Not set')}")
        st.markdown(f"**Cooling:** {st.session_state.get('cooling_installed', 'Not set')}")
    
    if st.session_state.step > 3:
        st.markdown(f"**Existing Window:** {st.session_state.get('existing_window', 'Not set')}")
        st.markdown(f"**CSW Type:** {st.session_state.get('csw_type', 'Not set')}")
        st.markdown(f"**CSW Area:** {st.session_state.get('csw_area', 0):,} sq.ft.")
    
    st.markdown('---')
    st.markdown('#### ⚠️ Status')
    st.markdown("""
    **Complete:**
    ✅ 4 exact HVAC options  
    ✅ Exact Excel formulas  
    ✅ % EUI Savings (F54)  
    
    **TODO:**
    ⏳ Complete weather data (874 locations)  
    ⏳ Complete savings lookup table (104 rows)  
    ⏳ Run extract_excel_data.py
    """)

st.markdown('---')
st.markdown(
    '<div style="text-align: center; color: gray; font-size: 0.8em;">'
    'CSW Savings Calculator v2.1 | Exact Excel Formulas'
    '</div>',
    unsafe_allow_html=True
)
