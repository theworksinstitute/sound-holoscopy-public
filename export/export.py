from config import ExportConfig
from export import wav_export, fft_export
from export.sheet_export import sheet_export
from signal_processing import signals
from signal_processing.signals import SignalRecording


def create_export(data: ExportConfig, log=print):
    """Main function that does exporting. You can pass custom log function, for example
    write the message to a text widget along with printing. It strings together all the other functions
    that export the sheets, audio and fft graphs.

    :param data: ExportConfig object that contains all the necessary data for exporting.
    :param log: function that takes a string and prints it somewhere.
    """

    log("Export initiated!")
    log("Reading signals...")
    # Read first signal and set fractions
    c1_signal: SignalRecording = signals.SignalRecording(data.c_files)
    c1_signal.read_files()

    # Read second signal and set fractions
    c2_signal: SignalRecording = signals.SignalRecording(data.ref_files)
    c2_signal.read_files()

    # Create sums and differences
    log("Creating sums and differences...")
    s1_sums, s2_sums, sd1, sd2 = signals.create_signal_combinations(c1_signal, c2_signal)

    sheet_export.create_export(s1_sums, s2_sums, sd1, sd2, data, log)

    if data.export_audio:
        wav_export.create_export(s1_sums, s2_sums, sd1, data, log)

    if data.export_fft:
        fft_export.create_export(s1_sums, s2_sums, sd1, data, log)

    log("Export complete!")
