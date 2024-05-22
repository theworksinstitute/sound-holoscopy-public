import numpy as np
import plotly.graph_objects as go
import os

from config import ExportConfig
from signal_processing import fft
from signal_processing.signals import Signal


def create_export(s1_sums, s2_sums, sd1, data: ExportConfig, log=print):
    log('Exporting FFT graphs...')

    signal_sets = data.get_signal_sets()
    fractions = data.time_fractions
    regions = data.frequency_regions
    x_lims = (min(regions, key=lambda x: x[0])[0], max(regions, key=lambda x: x[1])[1])
    x_lims = (x_lims[0] * 0.8, x_lims[1] * 1.2)  # Expand x_lims by 20% on each side

    for t in range(fractions):
        log(f'fraction {t + 1}/{fractions}')
        for i in range(len(signal_sets)):
            c1_signal: Signal = s1_sums[i].get_interval_fraction(fractions, t)
            c2_signal: Signal = s2_sums[i].get_interval_fraction(fractions, t)
            diffs_signal: Signal = sd1[i].get_interval_fraction(fractions, t)

            mics_str = 'mics_' + ''.join([str(m) for m in signal_sets[i]])
            time_str = f'fraction_{t + 1}_of_{fractions}'

            export_fft_fraction(data.fractions_path(), data.c_name,
                                f'{data.c_name}_{mics_str}_{time_str}.png', c1_signal, x_lims, log=log)
            export_fft_fraction(data.fractions_path(), data.ref_name,
                                f'{data.ref_name}_{mics_str}_{time_str}.png', c2_signal, x_lims, log=log)
            export_fft_fraction(data.fractions_path(), 'DIFF',
                                f'DIFF_{mics_str}_{time_str}.png', diffs_signal, x_lims, log=log)


def export_fft_fraction(audio_path, folder_name, file_name, signal: Signal, x_lims, log=print):
    """
    Exports the fft graph to the given path. Creates the directory if it doesn't exist.

    :param audio_path: path to the audio folder.
    :param folder_name: name of the folder to export to.
    :param file_name: name of the file to export to.
    :param signal: signal to export.
    :param x_lims: tuple of the x-axis limits.
    :param log: function that takes a string and prints it somewhere.
    """
    try:
        os.makedirs(f'{audio_path}/FFT/{folder_name}')
    except FileExistsError:
        pass

    log(f'exporting fft {file_name}')
    fft_data = fft.create_fft(signal)

    # Define logarithmic frequency range
    log_min_freq = np.log10(x_lims[0])
    log_max_freq = np.log10(x_lims[1])
    num_points = fft_data.frequency.shape[0] // 10  # Down-sample to every 10th point
    log_freq_range = np.logspace(log_min_freq, log_max_freq, num_points)

    # Find the nearest data points in the original frequency data efficiently
    nearest_indices = np.searchsorted(fft_data.frequency, log_freq_range)

    # Ensure the indices are within bounds
    nearest_indices = np.clip(nearest_indices, 0, len(fft_data.frequency) - 1)

    x = fft_data.frequency[nearest_indices]
    y = fft_data.plot[nearest_indices]
    min_y, max_y = np.min(y), np.max(y)

    # Create a Plotly figure with a logarithmic x-axis
    fig = go.Figure(data=go.Scatter(x=x, y=y, mode='lines', line=dict(width=1.5)),
                    layout_yaxis_range=[min_y - 6, max_y + 6])
    fig.update_xaxes(type='log')

    # Set labels and title
    fig.update_layout(
        title='FFT Data',
        xaxis_title='Frequency (Hz)',
        yaxis_title='Magnitude (dB)'
    )

    # Save the figure as a PNG file
    fig.write_image(f'{audio_path}/FFT/{folder_name}/{file_name}', width=800, height=600)
