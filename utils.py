# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from pathlib import Path

# Set visualization style
sns.set(style="whitegrid")
plt.rcParams['figure.figsize'] = (12, 8)

# Function to load all course data
def load_course_data():
    """Load all course data from CSV files"""
    course_data = {}
    
    # Path to the course ID tables
    course_id_path = Path('data/course_id_tables')
    
    # Load each CSV file
    for csv_file in course_id_path.glob('lp_courses_*.csv'):
        year = csv_file.stem.split('_')[-1]  # Extract year from filename
        df = pd.read_csv(csv_file)
        course_data[year] = df
        
    return course_data

# Function to load student preference data
def load_student_preferences():
    """Load student course preference data from Excel file"""
    try:
        # Try to load the student preference data
        preference_path = Path('data/student_data/student_course_preferences.xlsx')
        preferences_df = pd.read_excel(preference_path)
        return preferences_df
    except Exception as e:
        print(f"Could not load student preferences: {e}")
        # Create dummy data for demonstration
        return None


# Function to analyze course popularity
def analyze_course_popularity(preferences_df, courses_df):
    """
    Analyze course popularity based on student preferences.
    
    Parameters:
    - preferences_df: DataFrame containing student preferences with columns 'kurzXY'
      where X is the preference rank (1-4) and Y is a sequence number
    - courses_df: DataFrame containing course information with CourseID and name columns
    
    Returns:
    - DataFrame with popularity metrics for each course
    """
    # Get all preference columns (those starting with 'kurz')
    pref_columns = [col for col in preferences_df.columns if col.startswith('kurz')]
    
    # Create a dictionary to store course popularity data
    course_popularity = {}
    
    # Process each course ID
    for course_id in courses_df['CourseID'].astype(str):
        # Count how many times this course appears as 1st, 2nd, 3rd, and 4th preference
        pref_counts = {
            1: 0,  # 1st preference
            2: 0,  # 2nd preference
            3: 0,  # 3rd preference
            4: 0,  # 4th preference
        }
        
        for col in pref_columns:
            # Extract the preference rank (last digit of column name)
            try:
                pref_rank = int(col[-1])
                if pref_rank in pref_counts:
                    # Check if the column contains this course ID
                    course_count = (preferences_df[col] == float(course_id)).sum()
                    pref_counts[pref_rank] += course_count
            except ValueError:
                # Skip columns that don't end with a number 1-4
                continue
        
        # Calculate weighted score (higher weight for higher preferences)
        weighted_score = 4*pref_counts[1] + 3*pref_counts[2] + 2*pref_counts[3] + 1*pref_counts[4]
        total_mentions = sum(pref_counts.values())
        
        # Only include courses that were mentioned at least once
        if total_mentions > 0:
            # Get course name
            course_name = courses_df[courses_df['CourseID'] == int(course_id)]['name'].iloc[0]
            
            course_popularity[course_id] = {
                'course_id': int(course_id),
                'name': course_name,
                'pref_1': pref_counts[1],
                'pref_2': pref_counts[2],
                'pref_3': pref_counts[3],
                'pref_4': pref_counts[4],
                'total_mentions': total_mentions,
                'weighted_score': weighted_score,
                'avg_pref': weighted_score / total_mentions if total_mentions > 0 else 0
            }
    
    # Convert to DataFrame for easier analysis
    if course_popularity:
        popularity_df = pd.DataFrame.from_dict(course_popularity, orient='index')
        return popularity_df.sort_values('weighted_score', ascending=False)
    else:
        # Return empty DataFrame with expected columns
        return pd.DataFrame(columns=['course_id', 'name', 'pref_1', 'pref_2', 'pref_3', 'pref_4', 
                                     'total_mentions', 'weighted_score', 'avg_pref'])


# Function to visualize popularity
def plot_course_popularity(popularity_df, metric='weighted_score', top_n=10):
    """
    Plot top and bottom courses based on popularity metrics
    
    Parameters:
    - popularity_df: DataFrame from analyze_course_popularity containing course metrics
    - metric: Column to plot ('weighted_score', 'total_mentions', 'avg_pref', etc.)
    - top_n: Number of top/bottom courses to display
    """
    if popularity_df.empty:
        print("No data to visualize")
        return
        
    # Sort by the specified metric
    sorted_df = popularity_df.sort_values(by=metric, ascending=False)
    
    # Get top N and bottom N courses
    top_courses = sorted_df.head(top_n)
    bottom_courses = sorted_df.tail(top_n)
    
    # Create plot
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 16))
    
    # Plot top courses
    sns.barplot(x=metric, y='name', data=top_courses, ax=ax1, palette='viridis')
    ax1.set_title(f'Top {top_n} Courses by {metric}')
    ax1.set_xlabel(metric.replace('_', ' ').title())
    ax1.set_ylabel('Course Name')
    
    # Plot bottom courses
    sns.barplot(x=metric, y='name', data=bottom_courses, ax=ax2, palette='viridis')
    ax2.set_title(f'Bottom {top_n} Courses by {metric}')
    ax2.set_xlabel(metric.replace('_', ' ').title())
    ax2.set_ylabel('Course Name')
    
    plt.tight_layout()
    plt.show()


def save_popularity_data(popularity_df, year=None, output_dir=None):
    """
    Sorts the popularity DataFrame by weighted score and saves it to a CSV file.
    
    Parameters:
    - popularity_df: DataFrame containing course popularity metrics
    - year: String or int representing the year for the filename (optional)
    - output_dir: Directory to save the file (optional, default is current directory)
    
    Returns:
    - Path to the saved CSV file
    """
    if popularity_df.empty:
        print("Warning: Empty popularity DataFrame, nothing to save.")
        return None
    
    # Sort by weighted score (popularity) in descending order
    sorted_df = popularity_df.sort_values('weighted_score', ascending=False)
    
    # Create output directory if specified and doesn't exist
    if output_dir:
        from pathlib import Path
        output_path = Path(output_dir)
        output_path.mkdir(exist_ok=True, parents=True)
    else:
        output_path = Path('.')
    
    # Create filename
    year_suffix = f"_{year}" if year else ""
    filename = f"course_popularity_analysis{year_suffix}.csv"
    filepath = output_path / filename
    
    # Save to CSV
    sorted_df.to_csv(filepath)
    print(f"Popularity data saved to {filepath}")

