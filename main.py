import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# Load Excel file
data_file = './full hw.xlsx'
sheets = pd.ExcelFile(data_file)

# Load data into DataFrames
customers = sheets.parse('customers')
products = sheets.parse('products')
transaction_details = sheets.parse('transaction_details')
transactions = sheets.parse('transactions')

# ---- Task 1: Market Research ---- #
# Sales Trends
sales_trends = transactions.groupby('transaction_date')['total_amount'].sum()
sales_trends.index = pd.to_datetime(sales_trends.index)
sales_trends = sales_trends.resample('ME').sum()  # Use 'ME' instead of 'M'

# Plot Sales Trends
plt.figure(figsize=(10, 6))
plt.plot(sales_trends.index, sales_trends, marker='o', label='Monthly Sales')
plt.title('Monthly Sales Trends')
plt.xlabel('Date')
plt.ylabel('Total Sales')
plt.grid()
plt.legend()
plt.show()

# Customer Preferences
preferred_products = customers['preferred_product_id'].value_counts().rename_axis('product_id').reset_index(name='count')
preferred_products = preferred_products.merge(products, on='product_id', how='left')

# Plot Customer Preferences
plt.figure(figsize=(10, 6))
preferred_products.head(10).set_index('product_name')['count'].plot(kind='bar', color='skyblue')
plt.title('Top 10 Preferred Products by Customers')
plt.xlabel('Product Name')
plt.ylabel('Number of Customers')
plt.xticks(rotation=45, ha='right')
plt.show()

# Discounts and Sales Impact
discount_impact = transaction_details.groupby('discount')['total'].sum()

# Plot Discount Impact
plt.figure(figsize=(10, 6))
plt.plot(discount_impact.index, discount_impact, marker='o', label='Sales by Discount')
plt.title('Impact of Discounts on Sales')
plt.xlabel('Discount (%)')
plt.ylabel('Total Sales')
plt.grid()
plt.legend()
plt.show()

# ---- Task 2: Product Preference Analysis ---- #
# Preferences by Age and City
age_city_preferences = customers.groupby(['age', 'city'])['preferred_product_id'].value_counts().rename('count').reset_index()
age_city_preferences = age_city_preferences.merge(products, left_on='preferred_product_id', right_on='product_id', how='left')

# Plot Preferences for a Sample City and Age Group
sample_city = 'New York'  # Replace with a city of your choice
sample_age = 25  # Replace with an age of your choice
sample_preferences = age_city_preferences[(age_city_preferences['city'] == sample_city) & (age_city_preferences['age'] == sample_age)]

if not sample_preferences.empty:
    plt.figure(figsize=(10, 6))
    sample_preferences.set_index('product_name')['count'].plot(kind='bar', color='lightgreen')
    plt.title(f'Product Preferences for Age {sample_age} in {sample_city}')
    plt.xlabel('Product Name')
    plt.ylabel('Number of Customers')
    plt.xticks(rotation=45, ha='right')
    plt.show()
else:
    print(f"No data available for age {sample_age} in {sample_city}.")

# ---- Task 3: Customer Segmentation ---- #
# Segment Customers by Spending
customer_spending = transactions.groupby('customer_id')['total_amount'].sum().reset_index()
customer_spending = customer_spending.merge(customers[['customer_id', 'age', 'city']], on='customer_id', how='left')

# Define spending categories
bins = [0, 500, 1000, 5000, 10000, float('inf')]
labels = ['Low', 'Medium', 'High', 'Very High', 'Ultra']
customer_spending['spending_category'] = pd.cut(customer_spending['total_amount'], bins=bins, labels=labels)

# Visualize Spending Categories
spending_counts = customer_spending['spending_category'].value_counts()
plt.figure(figsize=(8, 6))
spending_counts.plot(kind='bar', color='coral')
plt.title('Customer Spending Categories')
plt.xlabel('Category')
plt.ylabel('Number of Customers')
plt.xticks(rotation=45)
plt.show()

# ---- Task 4: Sales Forecasting ---- #
# Prepare data for forecasting
forecast_data = sales_trends.copy()
forecast_data = forecast_data.reset_index()
forecast_data['Month'] = forecast_data['transaction_date'].dt.month
forecast_data['Year'] = forecast_data['transaction_date'].dt.year

# Simple forecast using previous trends
forecast_data['Forecast'] = forecast_data['total_amount'].shift(12)  # Seasonal trend

# Plot Forecast
plt.figure(figsize=(10, 6))
plt.plot(forecast_data['transaction_date'], forecast_data['total_amount'], label='Actual Sales', marker='o')
plt.plot(forecast_data['transaction_date'], forecast_data['Forecast'], label='Forecasted Sales', linestyle='--')
plt.title('Sales Forecasting')
plt.xlabel('Date')
plt.ylabel('Total Sales')
plt.grid()
plt.legend()
plt.show()

print("Analysis complete. Visualizations displayed.")
