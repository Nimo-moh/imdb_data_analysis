import pandas as pd

# Load the dataset
file_path = 'imdb_top_1000.xlsx'  # Update the path if necessary
data = pd.read_excel(file_path)

# Clean the Released_Year column
data['Released_Year'] = pd.to_numeric(data['Released_Year'], errors='coerce')
data = data.dropna(subset=['Released_Year'])
data['Released_Year'] = data['Released_Year'].astype(int)

# Ensure genre is split for movies with multiple genres
data['Genre'] = data['Genre'].str.split(',')

# Explode the genres for analysis
data_exploded = data.explode('Genre')

# Find the top-rated movie for each genre
top_rated_by_genre = data_exploded.loc[data_exploded.groupby('Genre')['IMDB_Rating'].idxmax()]
top_rated_by_genre = top_rated_by_genre[['Genre', 'Series_Title', 'IMDB_Rating']]

# Count movies released each year
release_trends = data['Released_Year'].value_counts().sort_index()
release_trends_df = release_trends.reset_index()
release_trends_df.columns = ['Release_Year', 'Movie_Count']

# Group by director and calculate average rating and movie count
popular_directors = data.groupby('Director').agg(
    Average_Rating=('IMDB_Rating', 'mean'),
    Movie_Count=('IMDB_Rating', 'count')
).sort_values(by=['Average_Rating', 'Movie_Count'], ascending=False)

# Save results as CSV files
top_rated_by_genre.to_csv('top_rated_by_genre.csv', index=False)
release_trends_df.to_csv('release_trends.csv', index=False)
popular_directors.to_csv('popular_directors.csv')

print("Analysis complete! CSV files saved successfully.")
