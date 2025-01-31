import duckdb
import pandas as pd

class VoltageStability:
    def __init__(self, dataset_path):
        """
        Initialize the VoltageStability class.

        Parameters:
        dataset_path (str): Path to the dataset file.
        """
        self.dataset_path = dataset_path
        self.df = None

    def get_dataset(self, scenario):
        """
        Load dataset filtered by a specific scenario.

        Parameters:
        scenario (str): The scenario to filter the dataset on.
        """
        query = f"""
        SELECT *
        FROM parquet_scan('{self.dataset_path}')
        WHERE scenario = '{scenario}'
        """
        try:
            self.df = duckdb.query(query).to_df()
            return self.df
        except Exception as e:
            raise ValueError(f"Failed to load dataset: {e}")

    def get_target_generator(self):
        """
        Analyze the dataset to determine voltage stability.

        Returns:
        dict: A dictionary containing the generator status and stability.
        """
        # Get generator out-of-step data
        out_of_step_col = [column for column in self.df.columns if 's:outofstep' in column]
        out_of_step_df = self.df[out_of_step_col].copy()
        out_of_step_df.columns = [column.split("_")[-1] for column in out_of_step_col]

        # Find collapsed generators
        out_of_step_sum = out_of_step_df.sum()
        generator_collapse = out_of_step_sum[out_of_step_sum != 0].index

        if len(generator_collapse) == 0:
            return {'generator': None, 'status': 'stable'}
        else:
            return {'generator': ', '.join(generator_collapse), 'status': 'unstable'}

# Example usage
# voltage_stability = VoltageStability('path/to/dataset.parquet')
# voltage_stability.get_dataset('scenario_name')
# result = voltage_stability.get_target()
# print(result)