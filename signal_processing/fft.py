from __future__ import annotations
import numpy as np
from numpy import ndarray

from signal_processing.signals import Signal


class FourierData:
    """
    Stores frequency (x) and fft value (y) and provides an interface for interacting with the data.
    """

    def __init__(self, frequency: ndarray, plot: ndarray):
        self.frequency = frequency
        self.plot = plot

    def get_region(self, start: int, end: int) -> FourierData:
        """
        :param start: Start index of the region.
        :param end: End index of the region.

        To get those indices, use corresponding methods from Signal class.
        """
        return FourierData(self.frequency[start:end], self.plot[start:end])

    def get_frequency_region(self, starting_frequency: int, ending_frequency: int) -> (int, int):
        start = np.where(self.frequency >= starting_frequency)[0][0]
        try:
            end = np.where(self.frequency >= ending_frequency)[0][0]
        except IndexError:
            end = self.frequency.size - 1
        return start, end

    def get_intervals(self, N: int) -> [(int, int)]:
        """
        :return: Start, end array indices for each interval. Can be then passed to get_region.
        :param N: How many intervals to divide fft into.
        """
        interval_length = int(len(self.frequency) / N)
        return [[i * interval_length, (i + 1) * interval_length] for i in range(N)]

    def get_interval(self, N: int, index: int):
        """
        :return: Start, end array indices for nth interval.
        :param N: How many intervals to divide the fft into.
        :param index: Which interval to get.
        """
        return self.get_intervals(N)[index]

    def get_max_amplitude(self):
        """
        :return: Frequency, amplitude for point where amplitude is max (across entire range).

        If you would like to get maximum amplitude on specific region, use get_region and call max_amplitude on the
        new object.
        """
        max_index = np.argmax(self.plot)
        max_frequency = self.frequency[max_index]
        max_amplitude = self.plot[max_index]
        return [max_frequency, max_amplitude]


def create_fft(signal: Signal, N=1, index=0) -> FourierData:
    """
    :param signal: Signal to create fft for.
    :param N: How many fractions to divide the signal into.
    :param index: Which index to get from N divisions.

    With default values (so passing only signal as an argument), fft will be created for whole signal.

    Resources:
    - [Power Spectrum in dB](https://gist.github.com/wiccy46/dc242c151858c48afe65cfa3bf935d20) Found in source
    code of someone's program.
    - [FFT to spectrum in decibel
    (dsp stackexchange)](https://dsp.stackexchange.com/questions/32076/fft-to-spectrum-in-decibel)
    - [How to compute dBFS?
    (dsp stackexchange)](https://dsp.stackexchange.com/questions/8785/how-to-compute-dbfs)
    - [Realtime_PyAudio_FFT/fft.py
        (GitHub)](https://github.com/aiXander/Realtime_PyAudio_FFT/blob/ef488e549780fee50f3012f28d6700a4a012e20c/src/fft.py)
    - [How to convert a .wav file to a spectrogram in python3
    (Stack Overflow)](https://stackoverflow.com/questions/44787437/how-to-convert-a-wav-file-to-a-spectrogram-in-python3)

    Just a side note, getting signal interval and then generating fft is different from getting fft from signal
    and then getting *fft* region.
    """

    signal_interval = signal.get_interval_fraction(N, index)
    data = signal_interval.data * np.hanning(signal_interval.length)
    FFT = np.fft.rfft(data, norm="forward")
    FFT = np.abs(FFT)
    x = np.fft.rfftfreq(len(data), 1 / signal_interval.samplerate)
    FFT = np.multiply(20, np.log10(FFT))

    return FourierData(x, FFT)
