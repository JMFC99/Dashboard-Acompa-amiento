import pandas as pd
import os
import argparse
import sys
import subprocess
import tempfile


def transform_columns(
    df, 
    id_column='Email', 
    transform_columns=None,
    column_type=bool,
    new_column_name='Tipo de Transformación'
):
    """
    Flexible transformation of columns with specific criteria
    
    Args:
        df (pandas.DataFrame): Input DataFrame
        id_column (str, optional): Column to use as identifier. Defaults to 'Email'.
        transform_columns (list, optional): Specific columns to transform. 
                                            If None, transforms all boolean columns.
        column_type (type, optional): Type of columns to transform. Defaults to bool.
        new_column_name (str, optional): Name for the new column with transformed values
    
    Returns:
        pandas.DataFrame: Transformed DataFrame
    """
    # If no specific columns provided, auto-detect columns of specified type
    if transform_columns is None:
        transform_columns = [
            col for col in df.columns 
            if df[col].dtype == column_type and col != id_column
        ]
    
    # Ensure the ID column is not in transform columns
    transform_columns = [col for col in transform_columns if col != id_column]
    
    # Validate inputs
    if not transform_columns:
        raise ValueError(f"No columns found of type {column_type}. Please specify transform_columns.")
    
    # Melt the DataFrame
    transformed_df = df.melt(
        id_vars=[id_column],  # Keep identifier column
        value_vars=transform_columns,
        var_name='Columna Original',
        value_name='Valor'
    )
    
    # Filter only rows where the value is True (or meets the specified condition)
    if column_type == bool:
        transformed_df = transformed_df[
            transformed_df['Valor'] == True
        ]
    else:
        # For non-boolean types, you might want to customize filtering
        transformed_df = transformed_df[
            transformed_df['Valor'].notna()  # Remove NaN values
        ]
    
    # Clean up the new column name
    transformed_df[new_column_name] = (
        transformed_df['Columna Original']
        .str.replace('Coaching ', '')  # Optional: remove specific prefixes
        .str.strip()
    )
    
    # Select and reorder columns
    result = transformed_df[[id_column, new_column_name]]
    
    return result


def process_excel_file(input_file, output_file, id_column='Email'):
    """
    Process Excel file with multiple sheets and transform data
    
    Args:
        input_file (str): Path to input Excel file
        output_file (str): Path to output Excel file
        id_column (str): Column to use as identifier
    """
    # Define sheet names to process
    sheet_names = [
        "Información básica", "Tipo de coaching",
        "Tipo de clientes", "Perfiles clientes",
        "Tipo industria", "ICF Certificación",
        "EMCC Certificación", "ICC Certificación",
        "WABC Certificación", "Assessments", "Otras certificaciones"
    ]
    
    print(f"Processing input file: {input_file}")
    print(f"Output will be saved to: {output_file}")
    
    # Validate input file exists
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    dfs = []
    
    # Process each sheet
    for i, sheet_name in enumerate(sheet_names):
        try:
            print(f"Processing sheet {i+1}/{len(sheet_names)}: {sheet_name}")
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            
            if i != 0:  # Transform all sheets except the first one
                transformed_df = transform_columns(
                    df, 
                    id_column=id_column, 
                    transform_columns=df.columns, 
                    new_column_name=sheet_name
                )
                dfs.append(transformed_df)
            else:  # Keep first sheet as is
                dfs.append(df)
                
        except Exception as e:
            print(f"Warning: Could not process sheet '{sheet_name}': {e}")
            continue
    
    if len(dfs) < 2:
        raise ValueError("Need at least 2 sheets to merge. Check your input file and sheet names.")
    
    # Merge all DataFrames
    print("Merging all sheets...")
    full_database = pd.merge(
        dfs[0], dfs[1], on=id_column, how='outer'
    )
    
    for i, df in enumerate(dfs):
        if i not in [0, 1]:
            full_database = pd.merge(
                full_database, df, on=id_column, how="outer"
            )
    
    # Create output directory if it doesn't exist
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Save the result
    full_database.to_excel(output_file, index=False)
    print(f"Successfully saved merged data to: {output_file}")
    print(f"Final dataset shape: {full_database.shape}")


def main():
    """Main function with argument parsing"""
    parser = argparse.ArgumentParser(
        description='Process Excel file with multiple sheets and merge coach data',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python script.py -i input.xlsx -o output.xlsx
  python script.py --input "data/coaches.xlsx" --output "results/dashboard.xlsx"
  python script.py -i input.xlsx -o output.xlsx --id-column "Email"
        '''
    )
    
    # Input file argument
    parser.add_argument(
        '-i', '--input',
        type=str,
        default='base_completa.xlsx',
        help='Path to input Excel file (default: input/2025-07-29/2025-07-29 Base de datos completa coaches.xlsx)'
    )
    
    # Output file argument
    parser.add_argument(
        '-o', '--output',
        type=str,
        default='2025-07-29-dashboard.xlsx',
        help='Path to output Excel file (default: 2025-07-29-dashboard.xlsx)'
    )
    
    # ID column argument
    parser.add_argument(
        '--id-column',
        type=str,
        default='Email',
        help='Column to use as identifier for merging (default: Email)'
    )
    
    # Verbose output
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='Enable verbose output'
    )
    
    # Parse arguments
    args = parser.parse_args()
    
    # Set verbose mode
    if not args.verbose:
        # Suppress pandas warnings if not in verbose mode
        import warnings
        warnings.filterwarnings('ignore')
    
    try:
        # Process the file
        process_excel_file(args.input, args.output, args.id_column)
        print("✅ Processing completed successfully!")
        
    except FileNotFoundError as e:
        print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
    except ValueError as e:
        print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()