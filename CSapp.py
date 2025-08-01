import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from scipy.stats import t
from statsmodels.stats.weightstats import DescrStatsW
import seaborn as sns
import matplotlib.pyplot as plt
import io

# Page configuration
st.set_page_config(
    page_title="Spreads & 12 Month Returns Analysis",
    page_icon="",
    layout="wide"
)

# Add company logo to top right
col1, col2, col3 = st.columns([6, 1, 1])
with col1:
    st.title("Credit Spreads vs 1Y Returns Analysis")
with col3:
    st.markdown('<div style="text-align: right; width: 100%; padding-right: 0; margin-right: 0; position: relative; right: 0;">', unsafe_allow_html=True)
    st.image("https://rubricsam.com/wp-content/uploads/2021/01/cropped-rubrics-logo-tight.png", width=350)
    st.markdown('</div>', unsafe_allow_html=True)

# Apply company color scheme and font
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Ringside:wght@300;400;500;600;700&display=swap');
    
    * {
        font-family: 'Ringside', sans-serif !important;
    }
    
    .stApp {
        background-color: #f8f9fa;
        font-family: 'Ringside', sans-serif !important;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        justify-content: space-between;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #D8D7DF;
        border-radius: 4px 4px 0px 0px;
        color: #001E4F;
        font-weight: 500;
        font-family: 'Ringside', sans-serif !important;
        min-width: 180px;
        flex: 1;
        text-align: center;
        padding: 8px 16px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2C5697;
        color: white;
    }
    .stButton > button {
        background-color: #2C5697;
        color: white;
        border: none;
        border-radius: 4px;
        padding: 8px 16px;
        font-family: 'Ringside', sans-serif !important;
    }
    .stButton > button:hover {
        background-color: #001E4F;
    }
    .stSelectbox > div > div {
        background-color: white;
        border: 1px solid #D8D7DF;
        font-family: 'Ringside', sans-serif !important;
    }
    .stSlider > div > div > div > div {
        background-color: #2C5697;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-family: 'Ringside', sans-serif !important;
    }
    
    h1 {
        color: #001E4F !important;
    }
    
    .stMarkdown {
        font-family: 'Ringside', sans-serif !important;
    }
    
    .stDataFrame {
        font-family: 'Ringside', sans-serif !important;
    }
    
    /* Rubrics Blue styling for multiselect buttons */
    .stMultiSelect [data-baseweb="select"] {
        background-color: #2C5697 !important;
        border-color: #2C5697 !important;
    }
    
    .stMultiSelect [data-baseweb="select"]:hover {
        background-color: #001E4F !important;
        border-color: #001E4F !important;
    }
    
    .stMultiSelect [data-baseweb="select"]:focus {
        background-color: #2C5697 !important;
        border-color: #2C5697 !important;
        box-shadow: 0 0 0 2px rgba(44, 86, 151, 0.2) !important;
    }
    
    /* Rubrics Blue styling for multiselect option buttons */
    .stMultiSelect [data-baseweb="tag"] {
        background-color: #2C5697 !important;
        border-color: #2C5697 !important;
        color: white !important;
    }
    
    .stMultiSelect [data-baseweb="tag"]:hover {
        background-color: #001E4F !important;
        border-color: #001E4F !important;
    }
    
    /* Rubrics Blue styling for multiselect dropdown */
    .stMultiSelect [data-baseweb="popover"] {
        border-color: #2C5697 !important;
    }
    
    .stMultiSelect [data-baseweb="option"]:hover {
        background-color: #2C5697 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Description
st.markdown("Upload your Excel file to analyze excess returns by spread categories and fixed income categories.")

# File upload
uploaded_file = st.file_uploader(
    "Choose an Excel file",
    type=['xlsx', 'xls'],
    help="Upload an Excel file with 'Excess Return' sheet containing spread and return data"
)

@st.cache_data
def load_and_process_data(file):
    """Load and process the uploaded Excel file"""
    try:
        spreadsheet = pd.ExcelFile(file)
        data = spreadsheet.parse(sheet_name='Excess Return')
        
        # Prepare and clean data for categorization
        categories = []
        for i in range(0, data.shape[1], 3):
            sub_df = data.iloc[:, i:i+3].dropna()
            sub_df.columns = ['Date', 'Spread', '1 Yr Ahead ER']
            category_name = data.columns[i+1].replace(' OAS', '').strip()
            sub_df['Category'] = category_name
            categories.append(sub_df)

        combined_df = pd.concat(categories, ignore_index=True)
        
        # Ensure Date column is datetime
        combined_df['Date'] = pd.to_datetime(combined_df['Date'])
        
        return combined_df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def categorize_spread(level):
    """Categorize spread levels into bins"""
    level_bps = level * 100  # convert from % to bps
    if level_bps < 100:
        return '<100'
    elif 100 <= level_bps < 150:
        return '100-150'
    elif 150 <= level_bps < 200:
        return '150-200'
    elif 200 <= level_bps < 250:
        return '200-250'
    elif 250 <= level_bps < 300:
        return '250-300'
    elif 300 <= level_bps < 400:
        return '300-400'
    elif 400 <= level_bps < 600:
        return '400-600'
    elif 600 <= level_bps < 800:
        return '600-800'
    else:
        return '800+'

def create_confidence_interval_plot(combined_df):
    """Create the confidence interval plot"""
    # Apply categorization
    combined_df['Spread Category'] = combined_df['Spread'].apply(categorize_spread)
    
    # Define desired order
    category_order = ['<100', '100-150', '150-200', '200-250', '250-300', '300-400', '400-600', '600-800', '800+']
    
    fig = go.Figure()
    colors = px.colors.qualitative.Set1

    for idx, cat in enumerate(combined_df['Category'].unique()):
        confidence_intervals = []
        for spread_cat in category_order:
            subset = combined_df[(combined_df['Category'] == cat) & (combined_df['Spread Category'] == spread_cat)]['1 Yr Ahead ER']
            if len(subset) > 1:
                stats = DescrStatsW(subset)
                ci_low, ci_high = stats.tconfint_mean(alpha=0.05)
                mean = stats.mean
                confidence_intervals.append({
                    'Spread Category': spread_cat,
                    'Mean': mean,
                    'CI Lower': ci_low,
                    'CI Upper': ci_high,
                })

        if confidence_intervals:
            confidence_df = pd.DataFrame(confidence_intervals)
            color = colors[idx % len(colors)]

            fig.add_trace(go.Scatter(
                x=confidence_df['Spread Category'],
                y=confidence_df['Mean'],
                mode='lines+markers',
                name=f'{cat} Mean',
                line=dict(color=color, width=2),
                marker=dict(size=8),
                hoverinfo='text',
                hovertext=[f"{cat} Mean ({scat}): {mean:.2f}%" for scat, mean in zip(confidence_df['Spread Category'], confidence_df['Mean'])],
                opacity=0.7,
                legendgroup=cat
            ))

            fig.add_trace(go.Scatter(
                x=confidence_df['Spread Category'],
                y=confidence_df['CI Upper'],
                mode='lines+markers',
                name=f'{cat} CI Upper',
                line=dict(color=color, width=1, dash='dot'),
                marker=dict(size=6),
                hoverinfo='text',
                hovertext=[f"{cat} CI Upper ({scat}): {ci_upper:.2f}%" for scat, ci_upper in zip(confidence_df['Spread Category'], confidence_df['CI Upper'])],
                opacity=0.7,
                legendgroup=cat,
                showlegend=False
            ))

            fig.add_trace(go.Scatter(
                x=confidence_df['Spread Category'],
                y=confidence_df['CI Lower'],
                mode='lines+markers',
                name=f'{cat} CI Lower',
                line=dict(color=color, width=1, dash='dot'),
                marker=dict(size=6),
                hoverinfo='text',
                hovertext=[f"{cat} CI Lower ({scat}): {ci_lower:.2f}%" for scat, ci_lower in zip(confidence_df['Spread Category'], confidence_df['CI Lower'])],
                opacity=0.7,
                legendgroup=cat,
                showlegend=False
            ))

    fig.update_layout(
        title='Interactive Plot: Mean and 95% Confidence Interval of Excess Return by Spread Category',
        xaxis_title='Spread Category',
        yaxis_title='Excess Return (%)',
        hovermode='closest',
        legend=dict(title='Fixed Income Categories', itemclick='toggle', itemdoubleclick='toggleothers'),
        height=600
    )
    
    return fig

def create_summary_table(combined_df):
    """Create summary statistics table"""
    combined_df['Spread Category'] = combined_df['Spread'].apply(categorize_spread)
    
    summary_list = []
    category_order = ['<100', '100-150', '150-200', '200-250', '250-300', '300-400', '400-600', '600-800', '800+']

    for cat in combined_df['Category'].unique():
        for spread_cat in category_order:
            subset = combined_df[
                (combined_df['Category'] == cat) &
                (combined_df['Spread Category'] == spread_cat)
            ]['1 Yr Ahead ER']

            if len(subset) > 1:
                stats = DescrStatsW(subset)
                mean = stats.mean
                ci_low, ci_high = stats.tconfint_mean(alpha=0.05)
                count = len(subset)

                summary_list.append({
                    'Category': cat,
                    'Spread Category': spread_cat,
                    'CI Lower': round(ci_low, 2),
                    'Mean': round(mean, 2),
                    'CI Upper': round(ci_high, 2),
                    'N': count
                })

    summary_df = pd.DataFrame(summary_list)
    
    # Sort to match desired format
    summary_df['Spread Category'] = pd.Categorical(
        summary_df['Spread Category'],
        categories=category_order,
        ordered=True
    )
    summary_df.sort_values(by=['Category', 'Spread Category'], inplace=True)
    
    return summary_df

def create_heatmap(combined_df):
    """Create heatmap visualization with date range selection and positive=green, negative=red"""
    # Add date range selector
    min_date = combined_df['Date'].min()
    max_date = combined_df['Date'].max()
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Heatmap Start Date", value=min_date, min_value=min_date, max_value=max_date, key="heatmap_start")
    with col2:
        end_date = st.date_input("Heatmap End Date", value=max_date, min_value=min_date, max_value=max_date, key="heatmap_end")
    
    # Filter data for heatmap
    filtered_df = combined_df[
        (combined_df['Date'] >= pd.Timestamp(start_date)) &
        (combined_df['Date'] <= pd.Timestamp(end_date))
    ]
    
    # Create summary table from filtered data
    filtered_df['Spread Category'] = filtered_df['Spread'].apply(categorize_spread)
    
    summary_list = []
    category_order = ['<100', '100-150', '150-200', '200-250', '250-300', '300-400', '400-600', '600-800', '800+']

    for cat in filtered_df['Category'].unique():
        for spread_cat in category_order:
            subset = filtered_df[
                (filtered_df['Category'] == cat) &
                (filtered_df['Spread Category'] == spread_cat)
            ]['1 Yr Ahead ER']

            if len(subset) > 1:
                stats = DescrStatsW(subset)
                mean = stats.mean
                ci_low, ci_high = stats.tconfint_mean(alpha=0.05)
                count = len(subset)

                summary_list.append({
                    'Category': cat,
                    'Spread Category': spread_cat,
                    'CI Lower': round(ci_low, 2),
                    'Mean': round(mean, 2),
                    'CI Upper': round(ci_high, 2),
                    'N': count
                })

    summary_df = pd.DataFrame(summary_list)
    
    # Sort to match desired format
    summary_df['Spread Category'] = pd.Categorical(
        summary_df['Spread Category'],
        categories=category_order,
        ordered=True
    )
    summary_df.sort_values(by=['Category', 'Spread Category'], inplace=True)
    
    # Create pivot table for heatmap
    pivot_df = summary_df.pivot(index="Category", columns="Spread Category", values="Mean")
    
    # Define company color scheme (negative=orange, positive=blue)
    custom_colorscale = [
        [0.0, "#CF4520"],    # Rubrics Orange for strong negative
        [0.25, "#CF4520"],   # Rubrics Orange for moderate negative
        [0.5, "white"],      # Neutral
        [0.75, "#2C5697"],   # Rubrics Medium Blue for moderate positive
        [1.0, "#001E4F"]     # Rubrics Blue for strong positive
    ]
    
    # Plot heatmap explicitly with custom colorscale
    fig = px.imshow(
        pivot_df,
        text_auto=True,
        aspect="auto",
        title=f"Heatmap of Mean Excess Return by Spread Category and Asset Class ({start_date} to {end_date})",
        labels=dict(x="Spread Category", y="Asset Class", color="Mean Excess Return (%)"),
        color_continuous_scale=custom_colorscale,
        color_continuous_midpoint=0
    )
    
    fig.update_layout(height=900)
    return fig, pivot_df, summary_df

def value_to_color(val, vmin=-20, vmax=30):
    """Map a value to a color on the custom diverging scale (red-white-green)."""
    # Clamp value
    val = max(min(val, vmax), vmin)
    
    # Define color stops for the diverging scale
    if val <= 0:
        # Negative values: red to white
        intensity = abs(val) / abs(vmin) if vmin != 0 else 0
        intensity = min(intensity, 1.0)
        # Interpolate from white (0) to darkred (1)
        red = 1.0
        green = 1.0 - intensity * 0.8  # Keep some green for lighter reds
        blue = 1.0 - intensity * 0.8
    else:
        # Positive values: white to green
        intensity = val / vmax if vmax != 0 else 0
        intensity = min(intensity, 1.0)
        # Interpolate from white (0) to darkgreen (1)
        red = 1.0 - intensity * 0.8
        green = 1.0
        blue = 1.0 - intensity * 0.8
    
    return f'rgb({int(red*255)}, {int(green*255)}, {int(blue*255)})'

def calculate_min_safe_spreads(combined_df, safety_threshold=0.05, min_obs=10):
    """Calculate minimum safe spread thresholds for each category."""
    df = combined_df.copy()
    
    # Use existing spread categories instead of narrow bins for better observation counts
    df['Spread Category'] = df['Spread'].apply(categorize_spread)
    
    # Define spread category order from tightest to widest
    spread_order = ['<100', '100-150', '150-200', '200-250',
                    '250-300', '300-400', '400-600', '600-800', '800+']
    
    # Group and calculate negative return probability by spread category
    grouped_stats = df.groupby(['Category', 'Spread Category']).agg(
        avg_return=('1 Yr Ahead ER', 'mean'),
        std_return=('1 Yr Ahead ER', 'std'),
        percent_negative=('1 Yr Ahead ER', lambda x: (x < 0).mean()),
        observations=('1 Yr Ahead ER', 'count')
    ).reset_index()

    # Filter by minimum observations for statistical reliability
    reliable_bins = grouped_stats[grouped_stats['observations'] >= min_obs]

    results = []

    # For each category, sort from tightest to widest spread category and find the threshold
    for category in reliable_bins['Category'].unique():
        category_df = reliable_bins[reliable_bins['Category'] == category]
        
        # Sort by spread category order (tightest to widest)
        category_df['Spread_Category_Order'] = category_df['Spread Category'].map(
            {cat: i for i, cat in enumerate(spread_order)}
        )
        category_df = category_df.sort_values('Spread_Category_Order')

        # Find first category where the historical negative percentage drops to or below threshold
        safe_bins = category_df[category_df['percent_negative'] <= safety_threshold]

        if not safe_bins.empty:
            threshold_bin = safe_bins.iloc[0]  # first bin from tightest upward
            # Calculate negative return observations
            negative_obs = int(threshold_bin['percent_negative'] * threshold_bin['observations'])
            negative_percentage = (negative_obs / threshold_bin['observations']) * 100
            
            # Convert spread category to approximate basis points for display
            spread_category = threshold_bin['Spread Category']
            if spread_category == '<100':
                spread_bps = 50  # midpoint of <100
            elif spread_category == '800+':
                spread_bps = 900  # representative value for 800+
            else:
                # Extract numbers from category like "100-150" -> 125
                parts = spread_category.split('-')
                spread_bps = (int(parts[0]) + int(parts[1])) / 2
            
            results.append({
                'Category': category,
                'Min_Safe_Spread_Bps': spread_bps,
                'Spread_Category': spread_category,
                'Avg_Return_Pct': threshold_bin['avg_return'],  # Already in percentage format
                'Negative_Return_Observations': negative_obs,
                'Total_Observations': threshold_bin['observations'],
                'Historical_Negative_Rate': f"{negative_obs} / {threshold_bin['observations']} = {negative_percentage:.3f}%",
                'Volatility_Pct': threshold_bin['std_return']  # Already in percentage format
            })

    return pd.DataFrame(results)

def negative_return_probability_plot(combined_df, show_excluded=True, selected_categories=None, tolerance=8):
    df = combined_df.copy()
    df['Spread Category'] = df['Spread'].apply(categorize_spread)

    grouped_stats = df.groupby(['Category', 'Spread Category']).agg(
        percent_negative=('1 Yr Ahead ER', lambda x: (x < 0).mean() * 100),
        avg_return=('1 Yr Ahead ER', 'mean'),
        std_return=('1 Yr Ahead ER', 'std'),
        observations=('1 Yr Ahead ER', 'count')
    ).reset_index()

    spread_order = ['<100', '100-150', '150-200', '200-250', 
                    '250-300', '300-400', '400-600', '600-800', '800+']

    # Use provided selected_categories or default to all
    if selected_categories is None:
        selected_categories = sorted(grouped_stats['Category'].unique())

    # Define a clean, professional color palette
    color_palette = px.colors.qualitative.Set2
    color_map = {cat: color_palette[i % len(color_palette)] for i, cat in enumerate(sorted(grouped_stats['Category'].unique()))}

    fig = go.Figure()

    for i, category in enumerate(selected_categories):
        cat_df = grouped_stats[grouped_stats['Category'] == category]
        cat_df['Spread_Category_Order'] = cat_df['Spread Category'].apply(lambda x: spread_order.index(x))
        cat_df = cat_df.sort_values('Spread_Category_Order')

        x_vals = cat_df['Spread Category']
        y_vals = cat_df['percent_negative']
        customdata = cat_df[['avg_return', 'std_return', 'observations']].values

        # Split into investable vs excluded segments
        investable_y = [y if y <= tolerance else None for y in y_vals]
        excluded_y = [y if y > tolerance else None for y in y_vals]

        # Colored trace for investable
        fig.add_trace(go.Scatter(
            x=x_vals,
            y=investable_y,
            mode='lines+markers',
            name=category,
            customdata=customdata,
            line=dict(color=color_map[category]),
            marker=dict(color=color_map[category]),
            hovertemplate=(
                f"<b>{category}</b><br>"
                "Spread Category: %{x}<br>"
                "Negative Returns: %{y:.2f}%<br>"
                "Avg Return: %{customdata[0]:.2f}%<br>"
                "Volatility: %{customdata[1]:.2f}%<br>"
                "Observations: %{customdata[2]}<extra></extra>"
            )
        ))

        # Greyed-out trace for excluded
        if show_excluded:
            fig.add_trace(go.Scatter(
                x=x_vals,
                y=excluded_y,
                mode='lines+markers',
                name=f'{category} (excluded)',
                line=dict(color='lightgrey', dash='dot'),
                marker=dict(color='lightgrey'),
                hoverinfo='skip',
                showlegend=False
            ))

    # Add threshold line
    fig.add_hline(
        y=tolerance,
        line_dash="dash",
        line_color="red",
        annotation_text=f"Risk Tolerance: {tolerance}% of 1-year return observations are negative at this level",
        annotation_position="top left"
    )

    # Calculate y-axis range based on visible data only
    visible_data = []
    for category in selected_categories:
        cat_df = grouped_stats[grouped_stats['Category'] == category]
        cat_df['Spread_Category_Order'] = cat_df['Spread Category'].apply(lambda x: spread_order.index(x))
        cat_df = cat_df.sort_values('Spread_Category_Order')
        
        x_vals = cat_df['Spread Category']
        y_vals = cat_df['percent_negative']
        
        if show_excluded:
            # Include all data points when showing excluded ranges
            visible_data.extend(y_vals)
        else:
            # Only include data points that meet the tolerance (investable)
            investable_y = [y for y in y_vals if y <= tolerance]
            visible_data.extend(investable_y)
    
    # Set y-axis range based on visible data
    if visible_data:
        y_max = max(visible_data) + 2  # Add small buffer
        y_range = [0, max(y_max, tolerance + 2)]  # Ensure tolerance line is visible
    else:
        y_range = [0, tolerance + 5]  # Fallback if no visible data

    fig.update_layout(
        title="Historical Negative Return Probability by Spread Category",
        xaxis_title="Spread Category (bps)",
        yaxis_title="% Observations with Negative 1Y Returns",
        hovermode="closest",
        yaxis=dict(range=y_range),
        legend_title="Asset Category",
        height=600,
        xaxis=dict(
            categoryorder="array",
            categoryarray=spread_order
        )
    )

    return fig

def create_violin_plot(combined_df):
    """Create violin plot with date range selector and diverging color scale."""
    combined_df['Spread Category'] = combined_df['Spread'].apply(categorize_spread)
    
    spread_order = ['<100', '100-150', '150-200', '200-250',
                    '250-300', '300-400', '400-600', '600-800', '800+']

    # Date range selector
    min_date = combined_df['Date'].min()
    max_date = combined_df['Date'].max()
    
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start Date", value=min_date, min_value=min_date, max_value=max_date)
    with col2:
        end_date = st.date_input("End Date", value=max_date, min_value=min_date, max_value=max_date)
    
    # Category selector
    selected_category = st.selectbox("Select Category", ['All'] + list(combined_df['Category'].unique()))
    
    # Filter data
    filtered_df = combined_df[
        (combined_df['Date'] >= pd.Timestamp(start_date)) &
        (combined_df['Date'] <= pd.Timestamp(end_date))
    ]
    
    if selected_category != 'All':
        filtered_df = filtered_df[filtered_df['Category'] == selected_category]
    
    # Create violin plot with vertical color transitions
    fig = go.Figure()
    
    for spread_cat in spread_order:
        spread_data = filtered_df[filtered_df['Spread Category'] == spread_cat]['1 Yr Ahead ER']
        
        if not spread_data.empty:
            # Calculate mean for overall violin color
            mean_value = spread_data.mean()
            
            # Create color gradient based on mean value using company colors
            if mean_value <= 0:
                # Negative mean: Rubrics Orange gradient
                intensity = abs(mean_value) / 20  # Normalize to max negative
                intensity = min(intensity, 1.0)
                # Convert Rubrics Orange (#CF4520) to rgba with opacity
                color = f'rgba(207, 69, 32, {0.3 + intensity * 0.4})'  # #CF4520 with varying opacity
            else:
                # Positive mean: Rubrics Blue gradient
                intensity = mean_value / 30  # Normalize to max positive
                intensity = min(intensity, 1.0)
                # Convert Rubrics Medium Blue (#2C5697) to rgba with opacity
                color = f'rgba(44, 86, 151, {0.3 + intensity * 0.4})'  # #2C5697 with varying opacity
            
            fig.add_trace(go.Violin(
                y=spread_data,
                x=[spread_cat] * len(spread_data),
                name=spread_cat,
                box_visible=True,
                meanline_visible=True,
                spanmode='hard',
                legendgroup=spread_cat,
                scalegroup=spread_cat,
                line=dict(color='black'),
                fillcolor=color,
                opacity=0.3
            ))
    
    fig.update_layout(
        title=f"Excess Returns Distribution - {selected_category} ({start_date} to {end_date})",
        xaxis=dict(
            title="Spread Category",
            categoryorder="array",
            categoryarray=spread_order
        ),
        yaxis=dict(
            title="Excess Return (%)",
            zeroline=True,
            zerolinewidth=3,
            zerolinecolor='black'
        ),
        violingap=0.1,
        height=600
    )
    
    return fig

# Main app logic
if uploaded_file is not None:
    # Load and process data
    combined_df = load_and_process_data(uploaded_file)
    
    if combined_df is not None:
        st.success(f"Data loaded successfully! Shape: {combined_df.shape}")
        
        # Display basic info
        st.subheader("Data Overview")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Observations", len(combined_df))
        with col2:
            st.metric("Categories", combined_df['Category'].nunique())
        with col3:
            st.metric("Date Range", f"{combined_df['Date'].min().strftime('%Y-%m-%d')} to {combined_df['Date'].max().strftime('%Y-%m-%d')}")
        
        # Create tabs for different visualizations
        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "Heatmap", 
            "Average Returns by Spread",
            "Historical Risk Analysis",
            "Violin Plot",
            "Summary Table", 
            "Download Data"
        ])
        
        with tab1:
            st.subheader("Heatmap Analysis")
            fig_heatmap, pivot_df, heatmap_summary_df = create_heatmap(combined_df)
            st.plotly_chart(fig_heatmap, use_container_width=True)
            
            # Download button for heatmap data
            csv_heatmap = pivot_df.to_csv()
            st.download_button(
                label="Download Heatmap Data as CSV",
                data=csv_heatmap,
                file_name="heatmap_data.csv",
                mime="text/csv"
            )
        
        with tab2:
            st.subheader("Average Returns by Spread Level")
            
            # User explicitly selects reference sub-asset class
            reference_class = st.selectbox("Select Reference Sub-Asset Class", combined_df['Category'].unique())

            # Get the spread range for the selected reference class
            ref_class_data = combined_df[combined_df['Category'] == reference_class]
            min_spread_bps = ref_class_data['Spread'].min() * 100  # Convert decimal to basis points
            max_spread_bps = ref_class_data['Spread'].max() * 100  # Convert decimal to basis points
            
            # Display the valid range for the selected class in basis points
            st.info(f"**{reference_class}** spread range: {min_spread_bps:.0f} to {max_spread_bps:.0f} basis points")

            # User selects spread level using a slider
            input_spread_bps = st.slider(
                "Select Spread Level (basis points)", 
                min_value=int(min_spread_bps), 
                max_value=int(max_spread_bps), 
                value=int(min_spread_bps + (max_spread_bps - min_spread_bps) / 2), 
                step=1,
                help="Select a spread level. The analysis will find all dates when the reference class had spreads within ±10 basis points of this level."
            )

            # Define explicit spread range (±10 basis points around selected level)
            # Convert basis points to decimal for comparison with data
            lower_bound = (input_spread_bps - 10) / 100
            upper_bound = (input_spread_bps + 10) / 100

            # Find dates explicitly where the reference class spread is within the chosen range
            ref_dates = combined_df[(combined_df['Category'] == reference_class) &
                                    (combined_df['Spread'] >= lower_bound) &
                                    (combined_df['Spread'] <= upper_bound)]['Date']

            # Check if there are matching records
            if ref_dates.empty:
                st.warning(f"No historical data found for {reference_class} at the selected spread level of {input_spread_bps} basis points (±10 bps)")
                st.write(f"Available spread levels for {reference_class}:")
                available_spreads = sorted(ref_class_data['Spread'].unique() * 100)
                st.write(f"Min: {available_spreads[0]:.0f}, Max: {available_spreads[-1]:.0f} basis points")
            else:
                # Get explicitly corresponding data for all sub-asset classes on these dates
                filtered_df = combined_df[combined_df['Date'].isin(ref_dates)]

                # Display calculation methodology
                st.info(f"""
                **Calculation Methodology:**
                - Found {len(ref_dates)} dates when {reference_class} had spreads between {(input_spread_bps-10):.0f} and {(input_spread_bps+10):.0f} basis points
                - For each sub-asset class, calculated the average 1-year ahead excess return across these {len(ref_dates)} dates
                - This shows how other asset classes performed when {reference_class} was at the selected spread level
                """)

                # Explicitly calculate statistics per sub-asset class
                stats_df = filtered_df.groupby('Category')['1 Yr Ahead ER'].agg(['mean', 'std', 'max', 'min', 'count']).reset_index()
                stats_df.columns = ['Sub-Asset Class', 'Average Excess Return (%)', 'Std Deviation', 'Max Return', 'Min Return', 'Observations']

                # Explicitly display results
                st.subheader(f"Excess Returns Statistics when {reference_class} Spread ≈ {input_spread_bps} basis points (±10 bps)")
                st.dataframe(stats_df.round(2), use_container_width=True)

                # Create ordered bar chart with company colors
                # Sort by average excess return from lowest to highest
                sorted_stats = stats_df.sort_values('Average Excess Return (%)')
                
                # Create custom colors based on company palette
                colors = []
                for value in sorted_stats['Average Excess Return (%)']:
                    if value < 0:
                        # Negative values: use Rubrics Orange (#CF4520)
                        colors.append('#CF4520')
                    else:
                        # Positive values: use Rubrics Blue (#2C5697)
                        colors.append('#2C5697')
                
                # Create bar chart with custom colors
                fig = go.Figure(data=[
                    go.Bar(
                        x=sorted_stats['Sub-Asset Class'],
                        y=sorted_stats['Average Excess Return (%)'],
                        marker_color=colors,
                        text=sorted_stats['Average Excess Return (%)'].round(2),
                        textposition='auto',
                    )
                ])
                
                fig.update_layout(
                    title=f"Average Excess Returns by Sub-Asset Class (when {reference_class} Spread ≈ {input_spread_bps} bps)",
                    xaxis_title="Sub-Asset Class",
                    yaxis_title="Average Excess Return (%)",
                    height=500,
                    showlegend=False
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Download button for Average Returns by Spread data
                csv_avg_returns = stats_df.to_csv(index=False)
                st.download_button(
                    label="Download Average Returns by Spread Data as CSV",
                    data=csv_avg_returns,
                    file_name=f"average_returns_{reference_class}_{input_spread_bps}bps.csv",
                    mime="text/csv"
                )
        
        with tab3:
            st.subheader("Historical Risk Analysis by Spread Category")

            # Get all categories for the multiselect
            df_temp = combined_df.copy()
            df_temp['Spread Category'] = df_temp['Spread'].apply(categorize_spread)
            grouped_stats_temp = df_temp.groupby(['Category', 'Spread Category']).agg(
                percent_negative=('1 Yr Ahead ER', lambda x: (x < 0).mean() * 100),
                avg_return=('1 Yr Ahead ER', 'mean'),
                std_return=('1 Yr Ahead ER', 'std'),
                observations=('1 Yr Ahead ER', 'count')
            ).reset_index()
            
            all_categories = sorted(grouped_stats_temp['Category'].unique())

            # Controls
            show_excluded = st.checkbox("Show excluded spread ranges", value=True)
            
            selected_categories = st.multiselect(
                "Select asset categories to display:",
                options=all_categories,
                default=all_categories
            )
            
            tolerance = st.slider("Risk Tolerance (%)", min_value=0, max_value=80, value=8, step=5)

            # Create the negative return probability plot
            fig_negative_returns = negative_return_probability_plot(
                combined_df, 
                show_excluded=show_excluded, 
                selected_categories=selected_categories, 
                tolerance=tolerance
            )
            st.plotly_chart(fig_negative_returns, use_container_width=True)
            
            # Add explanation
            st.info("""
            ### How to Interpret This Plot:
            - **Lines**: Each line represents a different asset category
            - **X-axis**: Spread categories from tightest (<100 bps) to widest (800+ bps)
            - **Y-axis**: Percentage of historical observations with negative 1-year returns
            - **Red dashed line**: Your selected risk tolerance threshold
            - **Key insight**: Points below the red line meet your risk tolerance; points above exceed it
            
            **Interactive features:**
            - Hover over lines to see detailed statistics for each spread category
            - Adjust the risk tolerance slider to see how it affects the analysis
            - Categories with lines that stay below your tolerance line are safer at those spread levels
            """)
        
        with tab4:
            st.subheader("Distribution Analysis")
            fig_violin = create_violin_plot(combined_df)
            st.plotly_chart(fig_violin, use_container_width=True)
        
        with tab6:
            st.subheader("Download Processed Data")
            
            # Download original processed data
            csv_processed = combined_df.to_csv(index=False)
            st.download_button(
                label="Download Processed Data as CSV",
                data=csv_processed,
                file_name="processed_data.csv",
                mime="text/csv"
            )
            
            # Download as Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                combined_df.to_excel(writer, sheet_name='Processed Data', index=False)
                summary_df = create_summary_table(combined_df)
                summary_df.to_excel(writer, sheet_name='Summary Statistics', index=False)
                pivot_df = create_summary_table(combined_df).pivot(index="Category", columns="Spread Category", values="Mean")
                pivot_df.to_excel(writer, sheet_name='Heatmap Data')
            
            output.seek(0)
            st.download_button(
                label="Download All Data as Excel",
                data=output.getvalue(),
                file_name="spreads_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with tab5:
            st.subheader("Summary Statistics")
            summary_df = create_summary_table(combined_df)
            st.dataframe(summary_df, use_container_width=True)
            
            # Download button for summary
            csv = summary_df.to_csv(index=False)
            st.download_button(
                label="Download Summary as CSV",
                data=csv,
                file_name="spread_summary.csv",
                mime="text/csv"
            )

else:
    st.info("Please upload an Excel file to begin the analysis.")
    
    # Instructions
    st.markdown("""
    ### Instructions:
    1. **File Format**: Upload an Excel file (.xlsx or .xls)
    2. **Sheet Name**: The file should contain a sheet named 'Excess Return'
    3. **Data Structure**: Data should be organized in columns of 3 (Date, Spread, 1 Yr Ahead ER) for each category
    4. **Spread Format**: Spreads should be in decimal format (e.g., 0.015 for 150 bps)
    
    ### What you'll get:
    - **Heatmap**: Visual representation of mean returns across categories with date range selection
    - **Average Returns by Spread**: Analysis of cross-asset performance at specific spread levels
    - **Violin Plot**: Distribution analysis with date range selection
    - **Summary Table**: Statistical summary by category and spread level
    - **Minimum Safe Spreads**: Calculate minimum spread thresholds for desired safety levels
    - **Download Options**: Export results in CSV or Excel format
    """)

