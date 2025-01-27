import streamlit as st
import pandas as pd

def find_comparables(subject_property, dataset):
    """
    Finds the 5 most comparable properties for the given subject property,
    prioritizing those with the closest market value and VPU/VPR.

    Args:
        subject_property: A pandas Series representing the subject property.
        dataset: The pandas DataFrame containing all properties.

    Returns:
        A DataFrame with the 5 most comparable properties.
    """
    # Filter based on conditions
    filtered_df = dataset[
        (dataset['Hotel Name'] != subject_property['Hotel Name']) &
        (dataset['Property Address'] != subject_property['Property Address']) &
        (dataset['Owner Name/ LLC Name'] != subject_property['Owner Name/ LLC Name']) &
        (dataset['Owner Street Address'] != subject_property['Owner Street Address']) &
        (dataset['Hotel Class'] == subject_property['Hotel Class']) &
        (dataset['Type'] == 'Hotel') &
        (dataset['Market Value-2024'] >= subject_property['Market Value-2024'] - 100000) &
        (dataset['Market Value-2024'] <= subject_property['Market Value-2024'] + 100000) &
        (dataset['VPR'] <= subject_property['VPR'] * 1.5)
    ].copy()

    # Calculate differences
    filtered_df['Market_Value_Diff'] = abs(filtered_df['Market Value-2024'] - subject_property['Market Value-2024'])
    filtered_df['VPU_VPR_Diff'] = abs(filtered_df['VPR'] - subject_property['VPR'])
    filtered_df['Combined_Diff'] = filtered_df['Market_Value_Diff'] + filtered_df['VPU_VPR_Diff']

    # Sort and get the top 5
    filtered_df = filtered_df.sort_values(by=['Combined_Diff', 'Market_Value_Diff', 'VPU_VPR_Diff']).head(5)

    return filtered_df


def main():
    st.title("Hotel Comparable Analysis")

    # File upload
    uploaded_file = st.file_uploader("Upload your hotel data (CSV)", type="csv")

    if uploaded_file is not None:
        # Load data
        data = pd.read_csv(uploaded_file)

        # Select subject property
        subject_index = st.slider("Select Subject Property Index", min_value=0, max_value=len(data)-1, step=1)
        subject_property = data.iloc[subject_index]

        # Display subject property
        st.write("Subject Property:")
        st.dataframe(subject_property.to_frame().T)

        # Find comparables
        comparables = find_comparables(subject_property, data)

        # Display comparables
        st.write("Comparable Properties:")
        st.dataframe(comparables)


if __name__ == "__main__":
    main()

