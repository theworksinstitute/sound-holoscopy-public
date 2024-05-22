import os

from config import ExportConfig
from signal_processing import signals
from signal_processing.signals import Signal


def create_export(s1_sums, s2_sums, sd1, data: ExportConfig, log=print):
    log('Exporting wav files...')

    signal_sets = data.get_signal_sets()
    fractions = data.time_fractions

    for t in range(fractions):
        log(f'fraction {t + 1}/{fractions}')
        for i in range(len(signal_sets)):
            # Sums and differences are created in the same order, this is why we can index them this way
            c1_signal: Signal = s1_sums[i].get_interval_fraction(fractions, t)
            c2_signal: Signal = s2_sums[i].get_interval_fraction(fractions, t)
            diffs_signal: Signal = sd1[i].get_interval_fraction(fractions, t)

            mics_str = 'mics_' + ''.join([str(m) for m in signal_sets[i]])
            time_str = f'fraction_{t + 1}_of_{fractions}'

            export_audio_fraction(data.fractions_path(), data.c_name,
                                  f'{data.c_name}_{mics_str}_{time_str}.wav', c1_signal)
            export_audio_fraction(data.fractions_path(), data.ref_name,
                                  f'{data.ref_name}_{mics_str}_{time_str}.wav', c2_signal)
            export_audio_fraction(data.fractions_path(), 'DIFF', f'DIFF_{mics_str}_{time_str}.wav',
                                  diffs_signal)


def export_audio_fraction(audio_path, folder_name, file_name, signal, log=print):
    """
    Exports the audio file to the given path. Creates the directory if it doesn't exist.

    :param audio_path: path to the audio folder.
    :param folder_name: name of the folder to export to.
    :param file_name: name of the file to export to.
    :param signal: signal to export.
    :param log: function that takes a string and prints it somewhere.
    """
    try:
        os.makedirs(f'{audio_path}/WAV/{folder_name}')
    except FileExistsError:
        pass

    log(f'exporting audio {file_name}')
    signals.write_signal(f'{audio_path}/WAV/{folder_name}/{file_name}', signal)
