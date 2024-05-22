"""
Relevant Urls:
[scipy.io.wavfile.read — SciPy v1.9.1 Manual](https://docs.scipy.org/doc/scipy/reference/generated/scipy.io.wavfile.read.html)
[scipy.signal.hilbert - SciPy v1.11.2 Manual](https://docs.scipy.org/doc/scipy/reference/generated/scipy.signal.hilbert.html)
"""
from __future__ import annotations
from dataclasses import dataclass

import numpy as np
from scipy.signal import hilbert
from scipy.io import wavfile
from config import microphone_combinations as sets
from typing import Union, List


@dataclass
class Signal:
    """
    Functions for get fraction/interval work exactly the same way as for fft.
    """

    def __init__(self, samplerate, data):
        self.data = data
        self.samplerate = samplerate
        self.length_in_seconds = int(data.shape[0] / samplerate)
        self.length = data.shape[0]

    def _get_fraction(self, start: int, end: int) -> Signal:
        """
        :param start: Start index of returned fraction.
        :param end: End index of returned fraction.
        :return: Signal object that represents a single fraction.
        """
        return Signal(self.samplerate, self.data[start:end])

    def _get_intervals(self, N: int) -> [(int, int)]:
        """
        :return: List of intervals that represent different indices that slice the signal in N equal parts.
        :param N: Interval count to return.
        """
        # interval length
        i_l = int(self.length / N)
        return [[i * i_l, (i + 1) * i_l] for i in range(N)]

    def _get_interval(self, N: int, index: int) -> (int, int):
        """
        :return: Tuple that contains start and end indices of a single interval.
        :param N: Interval count to choose from.
        :param index: Index of the interval after slicing signal in N equal parts.
        """
        return self._get_intervals(N)[index]

    def get_interval_fraction(self, N: int, index: int) -> Signal:
        """
        :return: Signal object that is some fraction of the original signal
        :param N: Interval count to choose from.
        :param index: Index of the interval after slicing signal in N equal parts.
        """
        intervals = self._get_interval(N, index)
        return self._get_fraction(*intervals)


class SignalRecording:
    def __init__(self, file_name: Union[str, List[str]]):
        """
        This is a wrapper for iterating over several signals at once, and having a list of files until the program
        gets to the point to read them.
        """
        if type(file_name) == str:
            self.files = [f'{file_name} MIC{x}.wav' for x in range(1, 7)]
        else:
            self.files = file_name
        self.mic_signals: [Signal] = []

    def read_files(self):
        """
        Reads all the signals and saves them in a single list of signals.
        """
        self.mic_signals = [read_signal(file) for file in self.files]


def signal_sum(*signals: Signal) -> Signal:
    """
    This function sums any number of signals. Convenient when having a list of signals to sum.
    """
    data = sum([s.data for s in signals])
    return Signal(signals[0].samplerate, data)


def signal_diff(signal1: Signal, signal2: Signal) -> Signal:
    """
    Subtracts one signal from another.
    """
    return Signal(signal1.samplerate, signal1.data - signal2.data)


def read_signal(filename: str) -> Signal:
    """
    :return: Signal object using the file name.
    :param filename: Path to the audio file.
    """
    return Signal(*wavfile.read(filename))


def write_signal(filename: str, signal: Signal):
    """
    Writes a signal object into a .wav audio file.

    :param filename: Path to the audio file.
    :param signal: Signal to write.
    """
    wavfile.write(filename, data=signal.data, rate=signal.samplerate)


def create_signal_combinations(signal_s1: SignalRecording, signal_s2: SignalRecording) -> \
        [[Signal], [Signal], [Signal], [Signal]]:
    """
    Implements step 1, 2 and 3 from the notes.
    - Step 1 involves summing all the s1 and s2 signal combinations.
    - Step 2 involves subtracting summed s2 from summed s1.
    - Step 3 involves subtracting the hilbert-transformed s2 from s1.
    """
    s1_sums: [Signal] = []
    s2_sums: [Signal] = []
    diffs: [Signal] = []
    diffs_hilbert: [Signal] = []
    signal_sets = get_signal_sets(len(signal_s2.mic_signals))
    for signal_set in signal_sets:
        s1_signals = [signal_s1.mic_signals[i - 1] for i in signal_set]
        s1_sum = signal_sum(*s1_signals)

        s2_signals = [signal_s2.mic_signals[i - 1] for i in signal_set]
        s2_sum = signal_sum(*s2_signals)

        s1_sums.append(s1_sum)
        s2_sums.append(s2_sum)

        s1a = Signal(s1_sum.samplerate, np.imag(hilbert(s1_sum.data)))
        s2b = Signal(s2_sum.samplerate, np.real(hilbert(s2_sum.data)))
        diffs_hilbert.append(signal_diff(s1a, s2b))

        diffs.append(signal_diff(s1_sum, s2_sum))
    return s1_sums, s2_sums, diffs, diffs_hilbert


def get_signal_sets(set_size: int) -> [[int]]:
    """
    This function tries to return a set of combinations of all the microphones that are valid.
    Raises a ValueError if a microphone set is not available.

    :return: List of microphone combinations. For stereo, it would be [[1], [2], [1, 2]].
    :param set_size: Specifies the number of microphones. 1 for mono, 2 for stereo, etc.
    """
    if set_size not in sets.keys():
        raise ValueError("Set size must be one of the values in " + str(list(sets.keys())))
    else:
        return sets[set_size]
