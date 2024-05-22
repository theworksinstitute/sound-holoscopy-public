from dataclasses import dataclass, field, InitVar
import os
from datetime import datetime
import json
from typing import Optional, Union, Literal

# Define root directory as the directory of this file
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


@dataclass
class ExportConfig:
    """
    This class contains all the configuration variables for the export module. It gets initialized with default values
    and then loads the values from the json file if it exists. Also, UI can change the values of this class.
    """
    # Initialization Variables
    json_load_path: InitVar[Optional[str]] = 'selected_files.json'

    # Private fields
    _json_data: dict[str, Union[str, int, list]] = None
    _destination_folder: str = field(default_factory=lambda: os.path.expanduser(
        f"~/Downloads/Export_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}"))
    _force_surround: bool = False  # Force surround export even if there are only 2 microphones. Used in some tests.

    # Regular fields
    c_files: [str] = None
    ref_files: [str] = None

    # Fields with default values
    time_fractions: int = 10
    export_audio: bool = False
    export_fft: bool = False
    frequency_regions: list[(int, int)] = field(default_factory=lambda: [
        [72, 74],
        [219, 221],
        [442, 444],
        [878, 880]
    ])
    c_name: str = 'C1'
    ref_name: str = 'REF'

    def __post_init__(self, json_load_path):
        # Do not load anything if json_load_path is None
        if json_load_path is None:
            return

        # This will return json with default values
        loaded_json = self.get_json()

        # If saved data exists, overwrite current json_data values
        if os.path.isfile(json_load_path):
            with open(json_load_path, 'r') as file:
                loaded_data = json.loads(file.read())
                for i, (k, v) in enumerate(loaded_data.items()):
                    loaded_json[k] = v
                loaded_json['regions'] = [tuple(pair) for pair in loaded_json['regions']]

        self._json_data = loaded_json

        # Update any variables with loaded json
        self.c_name = self._json_data['c_name']
        self.ref_name = self._json_data['ref_name']
        self.time_fractions = self._json_data['time_fractions']
        self.c_files = self._json_data['c']
        self.ref_files = self._json_data['ref']
        self.frequency_regions = self._json_data['regions']

    def get_signal_sets(self) -> [[int]]:
        return microphone_combinations[len(self.c_files)]

    def get_signal_sets_spatial(self) -> dict[str, [[int]]]:
        return microphone_combinations_spacial[len(self.c_files)]

    def get_spatial_mapping_to_signal_set(self, dim: Literal['x', 'y', 'z'] = 'x') -> [int]:
        """
        Returns a list of indices of signal sets that correspond to the given dimension.
        :param dim: Dimension to get the mapping for.
        :return: List of indices of signal sets that correspond to the given dimension. For example, if the dimension
        is 'x' and there are 2 microphones, then the returned list will be [0, 2, 1] because that's how the original
        signal sets are ordered in spatial combinations. Keep in mind that returned list will be less than or equal to
        the signal sets list. For example, if there are 4 microphones, then the returned list will have length of 5.
        """
        spatial_combinations = self.get_signal_sets_spatial()[dim]
        indices = []
        for comb in spatial_combinations:
            for i, s_set in enumerate(self.get_signal_sets()):
                if set(comb) == set(s_set):
                    indices.append(i)

        return indices

    def get_json(self):
        return {
            'c': self.c_files,
            'ref': self.ref_files,
            'c_name': self.c_name,
            'ref_name': self.ref_name,
            'time_fractions': self.time_fractions,
            'regions': self.frequency_regions,
        }

    def sheet_path(self) -> str:
        return os.path.join(self._destination_folder, f'{self._sheet_name()}.xlsx')

    def fractions_path(self) -> str:
        return os.path.join(self._destination_folder, 'Fractions')

    def get_mic_count(self):
        return len(self.c_files)

    def is_mono(self):
        return self.get_mic_count() == 1

    def is_surround(self):
        return self.get_mic_count() > 2 or self._force_surround

    def _sheet_name(self) -> str:
        return f'{self.c_name}_{self.ref_name}_{self.time_fractions}_time_fractions'


# Each key is a number of microphones, and each value is a list of
# microphone combinations. Usually is provided via ExportConfig class.
microphone_combinations = {
    1: ((1,),),
    2: (
        (1,),
        (2,),
        (1, 2)
    ),
    4: (
        (1,),
        (2,),
        (3,),
        (4,),
        (1, 2),
        (2, 3),
        (3, 4),
        (4, 1),
        (1, 2, 3),
        (2, 3, 4),
        (1, 3, 4),
        (1, 2, 4),
        (1, 2, 3, 4)
    ),
    6: (
        (1,),
        (2,),
        (3,),
        (4,),
        (5,),
        (6,),
        (1, 2),
        (2, 3),
        (3, 4),
        (4, 1),
        (1, 5),
        (5, 3),
        (3, 6),
        (6, 1),
        (2, 5),
        (5, 4),
        (4, 6),
        (6, 2),
        (1, 2, 5),
        (2, 3, 5),
        (3, 4, 5),
        (4, 1, 5),
        (1, 2, 6),
        (2, 3, 6),
        (3, 4, 6),
        (4, 1, 6),
        (1, 2, 3, 5),
        (2, 3, 4, 5),
        (3, 4, 1, 5),
        (4, 1, 2, 5),
        (1, 2, 3, 6),
        (2, 3, 4, 6),
        (3, 4, 1, 6),
        (4, 1, 2, 6),
        (5, 1, 6, 2),
        (5, 2, 6, 3),
        (5, 3, 6, 4),
        (5, 4, 6, 1),
        (1, 2, 3, 4, 5),
        (2, 3, 4, 5, 6),
        (3, 4, 5, 6, 1),
        (4, 5, 6, 1, 2),
        (5, 6, 1, 2, 3),
        (6, 1, 2, 3, 4),
        (1, 2, 3, 4, 5, 6),
    ),
}

microphone_combinations_spacial = {
    1:
        {
            'x': ((1,),)
        },
    2:
        {
            'x': ((1,), (1, 2), (2,))
        },
    4:
        {
            'x': ((4,), (1, 3, 4), (1, 2, 3, 4), (1, 2, 3), (2,)),
            'y': ((1,), (1, 2, 4), (1, 2, 3, 4), (2, 3, 4), (3,)),
        },
    6:
        {
            'x': ((3, 4), (3, 4, 5, 6), (1, 2, 3, 4, 5, 6), (1, 2, 5, 6), (1, 2)),
            'y': ((5,), (1, 4, 5), (1, 2, 3, 4, 5), (1, 2, 3, 4, 5, 6), (1, 2, 3, 4, 6), (2, 3, 6), (6,)),
            'z': ((1, 4, 6), (1, 2, 3, 4, 5, 6), (2, 3, 5))
        }
}
