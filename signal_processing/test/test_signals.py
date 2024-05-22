from signal_processing import signals
from signal_processing.signals import SignalRecording, create_signal_combinations
import numpy as np


def test__init():
	signal = SignalRecording('/path/audiofile')
	assert len(signal.files) == 6
	assert signal.files[1] == '/path/audiofile MIC2.wav'


def test_files():
	signal_ref: SignalRecording = SignalRecording('signal_processing/test/samples/input/E8_Test_REF1')
	signal_ref.read_files()
	signal_ref_mic1 = signal_ref.mic_signals[0]
	signal_ref_mic1_recording = signals.read_signal('signal_processing/test/samples/input/E8_Test_REF1 MIC1.wav')
	np.testing.assert_array_equal(signal_ref_mic1.data, signal_ref_mic1_recording.data)


def test_signal_set_differences():
	signal_ref: SignalRecording = SignalRecording('signal_processing/test/samples/input/E8_Test_REF1')
	signal_ref.read_files()
	signal_s: SignalRecording = SignalRecording('signal_processing/test/samples/input/E8_Test_S1')
	signal_s.read_files()
	[ref_sums, s_sums, diffs, _] = create_signal_combinations(signal_ref, signal_s)
	assert len(ref_sums) == 45
	assert len(s_sums) == 45
	assert len(diffs) == 45
	for [signal_sum, signal] in [[ref_sums, signal_ref], [s_sums, signal_s]]:
		assert signal_sum[0], signal.mic_signals[0]
		signal_sum12 = signals.signal_sum(signal.mic_signals[0], signal.mic_signals[1])
		signal_sum236 = signals.signal_sum(signal.mic_signals[1], signal.mic_signals[2], signal.mic_signals[5])
		assert signal_sum[6] == signal_sum12
		assert signal_sum[23] == signal_sum236
	signal_diff = signals.signal_diff(signal_s.mic_signals[0], signal_ref.mic_signals[0])
	assert diffs[0] == signal_diff
	ref_sum236 = signals.signal_sum(signal_ref.mic_signals[1], signal_ref.mic_signals[2], signal_ref.mic_signals[5])
	s_sum236 = signals.signal_sum(signal_s.mic_signals[1], signal_s.mic_signals[2], signal_s.mic_signals[5])
	diff236 = signals.signal_diff(s_sum236, ref_sum236)
	assert diffs[23] == diff236
