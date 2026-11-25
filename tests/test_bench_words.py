import random
import time

from bench.words import align, pages_lost, tokenize, word_scores


def make_words(count, seed=7, vocabulary=4000):
    rng = random.Random(seed)
    return [f"w{rng.randrange(vocabulary)}" for _ in range(count)]


def test_tokenize_folds_ligatures_and_case():
    assert tokenize("The ﬁnal FILE") == ["the", "final", "file"]


def test_tokenize_splits_cjk_per_ideograph_and_keeps_latin_words():
    assert tokenize("中文测试 abc, 日本語") == ["中", "文", "测", "试", "abc", "日", "本", "語"]


def test_tokenize_keeps_devanagari_word_with_vowel_signs_whole():
    word = "हिन्दी"
    assert tokenize(f"{word} {word}") == [word, word]


def test_tokenize_keeps_arabic_word_with_diacritics_whole():
    word = "كِتَاب"
    assert tokenize(f"x {word}") == ["x", word]


def test_tokenize_joiners_do_not_split_a_word():
    assert tokenize("a‍b b‌c") == ["ab", "bc"]


def test_tokenize_empty_and_punctuation_only():
    assert tokenize("") == []
    assert tokenize(" ... -- \t\n") == []


def test_align_identical_marks_everything():
    words = make_words(500)
    ref_flags, hyp_flags = align(words, list(words))
    assert sum(ref_flags) == 500 and sum(hyp_flags) == 500


def test_align_single_deletion_and_insertion_cost_one_word_each():
    ref = make_words(400)
    hyp = ref[:100] + ref[101:300] + ["zz_extra"] + ref[300:]
    ref_flags, hyp_flags = align(ref, hyp)
    assert sum(ref_flags) == 399
    assert sum(hyp_flags) == 399
    assert ref_flags[100] == 0


def test_align_empty_sides():
    assert sum(align([], ["a"])[1]) == 0
    assert sum(align(["a"], [])[0]) == 0


def test_word_scores_empty_reference_and_empty_hypothesis():
    assert word_scores([], [])["recall"] == 1.0
    assert word_scores(["a", "b", "c"], [])["recall"] == 0.0
    assert word_scores(["a", "b", "c"], [])["precision"] == 0.0
    assert word_scores([], ["a"])["precision"] == 0.0


def test_word_scores_report_order_loss_but_not_in_the_unordered_recall():
    ref = make_words(3000)
    moved = ref[1000:1300] + ref[:1000] + ref[1300:]
    scores = word_scores(ref, moved)
    assert scores["recall_unordered"] == 1.0
    assert scores["recall"] < 0.95


def test_word_scores_100k_identical_is_fast():
    words = make_words(100_000)
    started = time.perf_counter()
    scores = word_scores(words, list(words))
    assert scores["recall"] == 1.0 and scores["precision"] == 1.0
    assert time.perf_counter() - started < 10


def test_word_scores_100k_with_missing_and_moved_blocks_stays_bounded_and_order_aware():
    words = make_words(100_000)
    block = words[50_000:50_200]
    hyp = words[:10_000] + words[11_000:50_000] + words[50_200:80_000] + block + words[80_000:]
    started = time.perf_counter()
    scores = word_scores(words, hyp)
    assert time.perf_counter() - started < 20
    assert 0.980 < scores["recall"] < 0.992
    assert 0.985 < scores["recall_unordered"] < 0.995
    assert scores["recall_unordered"] > scores["recall"]


def test_pages_lost_counts_pages_whose_text_is_missing():
    pages = [make_words(60, seed=s) for s in (1, 2, 3)]
    hyp = pages[0] + pages[2]
    assert pages_lost(pages, hyp) == (3, 1)


def test_pages_lost_ignores_short_pages_and_reordered_pages():
    pages = [make_words(60, seed=1), ["a", "b", "c"], make_words(60, seed=3)]
    hyp = pages[2] + pages[0]
    assert pages_lost(pages, hyp) == (2, 0)


def test_pages_lost_is_zero_for_an_image_only_reference():
    assert pages_lost([[], []], ["x", "y", "z", "q", "r"]) == (0, 0)
