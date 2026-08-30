"""
Language detection module using Lingua AI + character-based detection.
"""
import re
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd


class LanguageChecker:
    """
    Language detector using hybrid approach:
    1. Character-based detection for definitive non-English scripts
    2. Lingua AI for ambiguous Latin-script text
    
    Performance optimizations:
    - Pre-compiled regex patterns
    - Vectorized pandas operations for batch processing
    - Cached detection results
    """
    
    # Non-English character patterns (definitive indicators)
    _NON_ENGLISH_PATTERNS = {
        'French': r'[àâçéèêëîïôùûœæ]',
        'German': r'[äöüß]',
        'Spanish': r'[ñ¿¡áéíóúü]',
        'Italian': r'[àèìòùé]',
        'Portuguese': r'[ãõçáéíóúâêô]',
        'Polish': r'[ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]',
        'Czech': r'[ěščřžýůúíďťňĚŠČŘŽÝŮÚÍĎŤŇ]',
        'Hungarian': r'[őűŐŰ]',
        'Turkish': r'[şğıİŞĞ]',
        'Romanian': r'[ăâîșțĂÂÎȘȚ]',
        'Russian': r'[а-яА-ЯёЁ]',
        'Ukrainian': r'[іїєґІЇЄҐ]',
        'Greek': r'[\u0370-\u03ff]',
        'Chinese': r'[\u4e00-\u9fff]',
        'Japanese': r'[\u3040-\u309f\u30a0-\u30ff\u4e00-\u9fff]',
        'Korean': r'[\uac00-\ud7af\u1100-\u11ff]',
        'Arabic': r'[\u0600-\u06ff]',
        'Hebrew': r'[\u0590-\u05ff]',
        'Thai': r'[\u0e00-\u0e7f]',
        'Vietnamese': r'[àảãáạăằẳẵắặâầẩẫấậèẻẽéẹêềểễếệìỉĩíịòỏõóọôồổỗốộơờởỡớợùủũúụưừửữứựỳỷỹýỵđ]',
    }
    
    # Pre-compiled regex patterns (class-level for performance)
    NON_ENGLISH_CHARS = {lang: re.compile(pattern) for lang, pattern in _NON_ENGLISH_PATTERNS.items()}
    
    # Common words/patterns that indicate specific languages (even without diacritics)
    # Used when special characters aren't present but word patterns clearly indicate language
    LANGUAGE_KEYWORDS = {
        'Portuguese': [
            r'\bnao\b', r'\bestao\b', r'\besta\b', r'\bsao\b', r'\bpor\b', r'\buma\b',
            r'\bmas\b', r'\bou\b', r'\bmigrando\b', r'\bobsolescencia\b', r'\bobsolecencia\b',
            r'\bsaldo\b', r'\bdiferencas\b', r'\blocalidade\b', r'\bunidade\b', r'\bdepartamento\b',
            r'\bprecisa\b', r'\bajuda\b', r'\bnome\b', r'\btelefone\b', r'\bpara\b',
            r'\bquando\b', r'\bonde\b', r'\bqual\b', r'\bquem\b', r'\bporque\b',
            # Additional Portuguese keywords from tickets
            r'\bconsigo\b', r'\bacessar\b', r'\bconsta\b', r'\btentar\b', r'\butilizar\b',
            r'\bnem\b', r'\bpois\b', r'\berro\b', r'\bao\b', r'\bcontato\b', r'\bseguros\b',
        ],
        'Spanish': [
            r'\bdiferencias\b', r'\berronea\b', r'\basignacion\b', r'\balmacen\b',
            r'\busuario\b', r'\breporta\b', r'\bdetecto\b', r'\bnumeros\b', r'\bconsignacion\b',
            r'\bquedando\b', r'\bmaterial\b', r'\bapoyar\b', r'\bnuevamente\b',
            r'\bhola\b', r'\bbuenos\b', r'\bdias\b', r'\bgracias\b', r'\bcomo\b',
            r'\bpuede\b', r'\bdar\b', r'\bbaja\b', r'\bpueden\b', r'\bsistema\b',
            # Additional common Spanish words
            r'\bproblemas\b', r'\bproblema\b', r'\beliminar\b', r'\bcapa\b', r'\bcuando\b',
            r'\brenuevo\b', r'\bcontrato\b', r'\baparece\b', r'\bnotifico\b', r'\bpermite\b',
            r'\brealizar\b', r'\berror\b', r'\ben\b', r'\bpara\b', r'\buna\b', r'\bno\b',
            r'\bsolicita\b', r'\bfavor\b', r'\bcuenta\b', r'\bcorreo\b', r'\bmensaje\b',
            r'\bdatos\b', r'\bnumero\b', r'\bequipo\b', r'\barchivo\b', r'\bfecha\b',
            r'\bnombre\b', r'\bcliente\b', r'\bsolicitud\b', r'\binformacion\b',
            r'\bpago\b', r'\bfactura\b', r'\bdonde\b', r'\bcual\b', r'\bquien\b',
            r'\btambien\b', r'\bahora\b', r'\besta\b', r'\bpero\b', r'\bsolo\b',
            # Additional Spanish keywords from tickets - verbs and common words
            r'\bmomento\b', r'\babrir\b', r'\baplicativo\b', r'\bingresar\b', r'\baplicacion\b',
            r'\barroja\b', r'\bdarle\b', r'\baparecen\b', r'\berrores\b', r'\bfigura\b',
            r'\btransmitir\b', r'\breversada\b', r'\brebookeada\b', r'\breverse\b', r'\bunas\b',
            r'\bduplico\b', r'\breversa\b', r'\bmercados\b', r'\bduplicados\b', r'\bduplicadas\b',
            r'\bcuotas\b', r'\bdetaladas\b', r'\bmercado\b', r'\bintentar\b', r'\bgrabar\b',
            r'\bcambios\b', r'\bfinalizar\b', r'\bsigue\b', r'\baparecciendo\b', r'\bestoy\b',
            r'\bintentando\b', r'\bsale\b', r'\bpantalla\b', r'\bnegra\b', r'\bcoinciden\b',
            r'\bmontos\b', r'\bingresados\b', r'\bcalculados\b', r'\bsolicito\b', r'\brecibe\b',
            r'\bincompleto\b', r'\bpuedo\b', r'\bingreso\b', r'\bnuevo\b', r'\bcolega\b',
            r'\bal\b', r'\bel\b', r'\blos\b', r'\blas\b', r'\bme\b', r'\bse\b', r'\bque\b',
        ],
        'French': [
            r'\bbonjour\b', r'\bcomment\b', r'\bmerci\b', r'\bpour\b', r'\bavec\b',
            r'\bsont\b', r'\bpas\b', r'\bune\b', r'\bles\b', r'\bvous\b', r'\bnous\b',
            r'\bqui\b', r"\bc'est\b", r'\bje\b', r'\btu\b', r'\bfonctionne\b',
            r'\bdepuis\b', r'\bmatin\b', r'\bprobleme\b', r'\bsysteme\b', r'\bfrancais\b',
            r'\bpourquoi\b', r'\bquand\b', r'\boui\b', r'\bnon\b', r'\baussi\b',
            r'\bmaintenant\b', r'\btoujours\b', r'\bpeut\b', r'\bfaire\b', r'\bdoit\b',
        ],
    }
    
    # Language name mappings (from Lingua names to display names)
    LANG_NAMES = {
        'FRENCH': 'French', 'GERMAN': 'German', 'SPANISH': 'Spanish', 'ITALIAN': 'Italian',
        'PORTUGUESE': 'Portuguese', 'DUTCH': 'Dutch', 'POLISH': 'Polish', 'RUSSIAN': 'Russian',
        'CHINESE': 'Chinese', 'JAPANESE': 'Japanese', 'KOREAN': 'Korean', 'ARABIC': 'Arabic',
        'SWEDISH': 'Swedish', 'DANISH': 'Danish', 'BOKMAL': 'Norwegian', 'NYNORSK': 'Norwegian',
        'CZECH': 'Czech', 'HUNGARIAN': 'Hungarian', 'TURKISH': 'Turkish', 'ROMANIAN': 'Romanian',
        'UKRAINIAN': 'Ukrainian', 'GREEK': 'Greek', 'HEBREW': 'Hebrew', 'THAI': 'Thai',
        'VIETNAMESE': 'Vietnamese', 'FINNISH': 'Finnish', 'SLOVAK': 'Slovak', 'CROATIAN': 'Croatian',
        'SLOVENE': 'Slovenian', 'BULGARIAN': 'Bulgarian', 'SERBIAN': 'Serbian', 'CATALAN': 'Catalan',
        # Character-based detection names (already display names)
        'French': 'French', 'German': 'German', 'Spanish': 'Spanish', 
        'Italian': 'Italian', 'Portuguese': 'Portuguese', 'Polish': 'Polish',
        'Czech': 'Czech', 'Russian': 'Russian', 'Chinese': 'Chinese',
        'Japanese': 'Japanese', 'Korean': 'Korean', 'Arabic': 'Arabic',
        'Hungarian': 'Hungarian', 'Turkish': 'Turkish', 'Romanian': 'Romanian',
        'Ukrainian': 'Ukrainian', 'Greek': 'Greek', 'Hebrew': 'Hebrew',
        'Thai': 'Thai', 'Vietnamese': 'Vietnamese',
    }

    # Which detection tier produced a hit. Reported per row so the UI can say
    # what actually found something rather than implying all three tiers ran.
    METHOD_CHARACTERS = 'characters'
    METHOD_KEYWORDS = 'keywords'
    METHOD_LINGUA = 'lingua'

    # Human-readable labels for the three detection modes offered in the UI.
    MODE_LABELS = {
        'all': 'characters + keywords + Lingua',
        'lingua_only': 'Lingua only',
        'rules_only': 'characters + keywords (no Lingua)',
    }
    
    # Pre-compiled keyword patterns for each language (class-level)
    _COMPILED_KEYWORDS = None
    
    # Pre-compiled text cleaning patterns
    _CLEAN_PATTERNS = [
        (re.compile(r'\b(INC|CHG|REQ|PRB|RITM|TASK|KB|SCTASK|SR|CR|WO|JIRA)\d+\b', re.IGNORECASE), ''),
        (re.compile(r'\b[A-Z]{2,}[0-9]+[A-Z0-9]*\b'), ''),
        (re.compile(r'\b[0-9]+[A-Z]+[A-Z0-9]*\b'), ''),
        (re.compile(r'\b[A-Z]{3,}[0-9A-Z]*\d+[A-Z0-9]*\b'), ''),
        (re.compile(r'\S+@\S+\.\S+'), ''),
        (re.compile(r'https?://\S+'), ''),
        (re.compile(r'www\.\S+'), ''),
        (re.compile(r'[A-Za-z]:\\[\w\\.\-]+'), ''),
        (re.compile(r'/[\w/.\-]+\.\w+'), ''),
        (re.compile(r'\[[A-Z0-9_\s\-]+\]', re.IGNORECASE), ''),
        (re.compile(r'\b\d{1,4}[-/]\d{1,2}[-/]\d{1,4}\b'), ''),
        (re.compile(r'\b\d+\b'), ''),
        (re.compile(r'\b(SAP|ERP|CRM|API|SQL|HTML|XML|JSON|PDF|CSV)\b', re.IGNORECASE), ''),
        (re.compile(r'\s+'), ' '),
    ]
    
    @classmethod
    def _get_compiled_keywords(cls):
        """Lazy-compile keyword patterns once."""
        if cls._COMPILED_KEYWORDS is None:
            cls._COMPILED_KEYWORDS = {
                lang: [re.compile(p, re.IGNORECASE) for p in patterns]
                for lang, patterns in cls.LANGUAGE_KEYWORDS.items()
            }
        return cls._COMPILED_KEYWORDS
    
    def __init__(self):
        self._lingua_detector = None
        self._lingua_available = None
        self._detection_cache = {}  # Cache for repeated text detection
    
    def _init_lingua(self) -> bool:
        """Initialize Lingua detector lazily."""
        if self._lingua_detector is not None:
            return True
        
        if self._lingua_available is False:
            return False
        
        try:
            from lingua import Language, LanguageDetectorBuilder
            self._lingua_detector = (
                LanguageDetectorBuilder
                .from_languages(
                    Language.ENGLISH, Language.FRENCH, Language.GERMAN,
                    Language.SPANISH, Language.ITALIAN, Language.PORTUGUESE,
                    Language.DUTCH, Language.POLISH, Language.RUSSIAN,
                    Language.CHINESE, Language.JAPANESE, Language.KOREAN,
                    Language.ARABIC, Language.SWEDISH, Language.DANISH,
                    Language.BOKMAL, Language.NYNORSK, Language.CZECH,
                    Language.HUNGARIAN, Language.TURKISH, Language.ROMANIAN,
                    Language.UKRAINIAN, Language.GREEK, Language.HEBREW,
                    Language.THAI, Language.VIETNAMESE, Language.FINNISH,
                    Language.SLOVAK, Language.CROATIAN, Language.SLOVENE,
                    Language.BULGARIAN, Language.SERBIAN, Language.CATALAN
                )
                .with_preloaded_language_models()
                .build()
            )
            self._lingua_available = True
            return True
        except ImportError:
            self._lingua_available = False
            return False
        except Exception:
            self._lingua_available = False
            return False
    
    def is_lingua_available(self) -> bool:
        """Check if Lingua is available."""
        if self._lingua_available is None:
            self._init_lingua()
        return self._lingua_available or False
    
    @classmethod
    def clean_text_for_detection(cls, text: str) -> str:
        """Remove noise that confuses language detection using pre-compiled patterns."""
        for pattern, replacement in cls._CLEAN_PATTERNS:
            text = pattern.sub(replacement, text)
        return text.strip()

    @classmethod
    def _apply_clean_patterns(cls, values: pd.Series) -> pd.Series:
        """Run the pattern list over a Series at the pandas level."""
        for pattern, replacement in cls._CLEAN_PATTERNS:
            values = values.str.replace(pattern, replacement, regex=True)
        return values.str.strip()

    @classmethod
    def clean_series(cls, series: pd.Series) -> pd.Series:
        """
        Vectorized equivalent of clean_text_for_detection for a whole Series.

        Cleaning is a pure function of the input string, so the patterns run
        over the *distinct* values and the result is mapped back. Ticket text
        repeats heavily — the same templated issue arrives hundreds of times,
        which is precisely what the recurring-issue check is looking for — so on
        a real export this collapses the work by one to two orders of magnitude.
        Each of the 14 patterns otherwise costs one re.sub per row: 100k rows
        over two description columns was 3.2M calls and ~5s.
        """
        text = series.astype(str)
        if text.empty:
            return cls._apply_clean_patterns(text)

        uniques = text.unique()
        # With no repetition the mapping cannot pay for itself, so clean directly.
        if len(uniques) == len(text):
            return cls._apply_clean_patterns(text)

        lookup = pd.Series(uniques, index=uniques, dtype=object)
        return text.map(cls._apply_clean_patterns(lookup))
    
    def detect_language(self, text: str, min_confidence: float = 0.50, 
                        mode: str = 'all') -> Optional[Tuple[str, float]]:
        """
        Detect language of a single text.
        
        Args:
            text: The text to analyze
            min_confidence: Minimum confidence threshold for Lingua detection (0.0-1.0)
            mode: Detection mode - 'all' (default), 'lingua_only', or 'rules_only'
                  - 'all': Use character + keyword + Lingua detection (recommended)
                  - 'lingua_only': Only use Lingua ML detection (slower but consistent)
                  - 'rules_only': Only use character + keyword detection (faster but limited)
        
        Returns: (language_name, confidence) or None if English/unknown
        """
        detailed = self.detect_language_detailed(text, min_confidence, mode)
        return (detailed[0], detailed[1]) if detailed else None

    def detect_language_detailed(self, text: str, min_confidence: float = 0.50,
                                 mode: str = 'all') -> Optional[Tuple[str, float, str]]:
        """
        Same detection as detect_language, but also reports which tier fired.

        Returns: (language_name, confidence, method) or None, where method is
        one of METHOD_CHARACTERS / METHOD_KEYWORDS / METHOD_LINGUA.

        The method matters because the three tiers are not comparable: a
        character hit is definitive (a Cyrillic string is not English), a
        keyword hit is a fixed-confidence heuristic, and only a Lingua hit
        carries a real model probability. Collapsing them into one "confidence"
        column made a 0.90 keyword guess look like a model score.
        """
        if not text or not isinstance(text, str):
            return None

        text = str(text).strip()
        if len(text) < 8:
            return None

        # Check cache first (significant speedup for repeated texts)
        cache_key = (text[:200], mode, min_confidence)  # Limit key size
        if cache_key in self._detection_cache:
            return self._detection_cache[cache_key]

        result = self._detect_language_uncached(text, min_confidence, mode)

        # Cache result (limit cache size to prevent memory bloat)
        if len(self._detection_cache) < 10000:
            self._detection_cache[cache_key] = result

        return result

    def _detect_language_uncached(self, text: str, min_confidence: float,
                                   mode: str) -> Optional[Tuple[str, float, str]]:
        """Internal detection without caching. Returns (lang, confidence, method)."""
        # Step 1: Check for definitive non-English characters (highest priority)
        # Skip if lingua_only mode
        if mode != 'lingua_only':
            # Use pre-compiled patterns
            for lang, pattern in self.NON_ENGLISH_CHARS.items():
                if pattern.search(text):
                    return (lang, 1.0, self.METHOD_CHARACTERS)

            # Step 2: Check keyword patterns using pre-compiled regex
            text_lower = text.lower()
            compiled_keywords = self._get_compiled_keywords()
            for lang, patterns in compiled_keywords.items():
                match_count = sum(1 for p in patterns if p.search(text_lower))
                if match_count >= 2:  # At least 2 keyword matches = strong signal
                    return (lang, 0.90, self.METHOD_KEYWORDS)

        # Step 3: Use Lingua on cleaned text for Latin-script languages
        # Skip if rules_only mode
        if mode != 'rules_only' and self._init_lingua():
            cleaned = self.clean_text_for_detection(text)
            # Lowered requirements for short texts
            if len(cleaned) >= 8 and len(cleaned.split()) >= 1:
                try:
                    result = self._lingua_detector.compute_language_confidence_values(cleaned)
                    if result:
                        top_lang = result[0]
                        lang_name = top_lang.language.name
                        conf = top_lang.value

                        if lang_name != 'ENGLISH' and conf > min_confidence:
                            return (self.LANG_NAMES.get(lang_name, lang_name), conf,
                                    self.METHOD_LINGUA)

                        # Fallback: if English barely wins, check 2nd place
                        if len(result) > 1 and lang_name == 'ENGLISH':
                            second = result[1]
                            if second.value > 0.35 and (conf - second.value) < 0.15:
                                return (self.LANG_NAMES.get(second.language.name, second.language.name),
                                        second.value, self.METHOD_LINGUA)
                except Exception:
                    pass

        return None
    
    def detect_language_for_row(self, text: str, min_confidence: float = 0.50,
                                 mode: str = 'all') -> str:
        """
        Detect language for a single row/text - returns language name or 'English'.
        Used for per-row language detection in exports.
        
        Args:
            text: The text to analyze
            min_confidence: Minimum confidence threshold for Lingua (0.0-1.0)
            mode: 'all', 'lingua_only', or 'rules_only'
        """
        result = self.detect_language(text, min_confidence, mode)
        if result:
            lang, _ = result
            return self.LANG_NAMES.get(lang, lang)
        return 'English'
    
    def check_column(self, df: pd.DataFrame, column: str, 
                     min_length: int = 10, min_confidence: float = 0.50,
                     mode: str = 'all') -> Dict[str, Any]:
        """
        Check a DataFrame column for non-English text.
        Optimized with vectorized filtering and batch processing.
        
        Returns: {
            'column': str,
            'total_analyzed': int,
            'total_original': int,
            'skipped_short': int,        # dropped for being under min_length
            'min_length': int,
            'non_english_count': int,
            'percentage': float,
            'languages': {lang: count},
            'methods': {method: count},  # which tier found each detection
            'samples': [(idx, text, lang, confidence, method)],
            'mode': str,
            'mode_label': str,           # human-readable description of the mode
            'rules_used': bool,          # character/keyword tiers actually ran
            'using_lingua': bool,        # Lingua actually ran for THIS call
            'lingua_available': bool,    # Lingua is installed at all
        }

        Note that 'using_lingua' means "Lingua contributed to this run", not
        "Lingua is installed" - those differ whenever mode is 'rules_only',
        and reporting the latter made the UI claim ML detection had run when
        it had been explicitly switched off.
        """
        if column not in df.columns:
            return {
                'column': column,
                'error': f'Column {column} not found',
                'non_english_count': 0
            }

        lingua_available = self.is_lingua_available()
        rules_used = mode != 'lingua_only'
        lingua_used = mode != 'rules_only' and lingua_available
        run_info = {
            'mode': mode,
            'mode_label': self.MODE_LABELS.get(mode, mode),
            'rules_used': rules_used,
            'using_lingua': lingua_used,
            'lingua_available': lingua_available,
            'min_length': min_length,
        }

        # Vectorized operations for filtering
        col_data = df[column]
        mask_notna = col_data.notna()
        original_count = mask_notna.sum()

        # Convert to string and filter by length in one pass
        text_series = col_data[mask_notna].astype(str)
        length_mask = text_series.str.len() >= min_length
        text_series = text_series[length_mask]
        total_count = len(text_series)

        empty_result = {
            'column': column,
            'total_analyzed': total_count,
            'total_original': int(original_count),
            'skipped_short': int(original_count) - total_count,
            'non_english_count': 0,
            'percentage': 0,
            'languages': {},
            'methods': {},
            'samples': [],
            **run_info,
        }

        # A check that could not run must never be reported as a check that
        # passed. With mode='lingua_only' and Lingua missing, every tier is
        # disabled, so detection returns None for every row - which previously
        # rendered as "all entries appear to be in English" on a column full
        # of French and German.
        if not rules_used and not lingua_used:
            return {
                **empty_result,
                'warning': (
                    "Lingua-only mode was selected, but the Lingua language "
                    "detector is not installed - so no detection ran and this "
                    "column has NOT been checked. Switch Mode to 'all' or "
                    "'rules_only', or install lingua-language-detector."
                ),
            }

        if total_count == 0:
            return {
                **empty_result,
                'warning': f'No text data to analyze (all entries under {min_length} characters)',
            }

        non_english_rows = []
        lang_counts = {}
        method_counts = {}

        # Process in batches for better memory efficiency
        batch_size = 1000
        texts_list = list(text_series.items())

        for i in range(0, len(texts_list), batch_size):
            batch = texts_list[i:i + batch_size]
            for idx, text in batch:
                result = self.detect_language_detailed(text, min_confidence, mode)
                if result:
                    detected_lang, confidence, method = result
                    non_english_rows.append((idx, text, detected_lang, confidence, method))
                    lang_counts[detected_lang] = lang_counts.get(detected_lang, 0) + 1
                    method_counts[method] = method_counts.get(method, 0) + 1

        non_english_count = len(non_english_rows)
        pct = (non_english_count / total_count) * 100 if total_count > 0 else 0

        # Get top samples sorted by confidence (use heapq for large lists)
        if len(non_english_rows) > 100:
            import heapq
            sorted_samples = heapq.nlargest(10, non_english_rows, key=lambda x: x[3])
        else:
            sorted_samples = sorted(non_english_rows, key=lambda x: -x[3])[:10]

        return {
            'column': column,
            'total_analyzed': total_count,
            'total_original': int(original_count),
            'skipped_short': int(original_count) - total_count,
            'non_english_count': non_english_count,
            'percentage': round(pct, 2),
            'languages': lang_counts,
            'methods': method_counts,
            'samples': sorted_samples,
            # The per-row answer, built from work already done rather than by a
            # second pass. Detection is the slowest thing in the app (~1.2ms per
            # distinct text with Lingua), and this loop has just decided every
            # row - returning only 10 samples threw that away, so a report
            # wanting a per-ticket language had no choice but to redo all of it.
            #
            # Only non-English rows are keyed, so the dict stays proportional to
            # the finding rather than to the file. No text is retained.
            'row_languages': {idx: lang for idx, _text, lang, _conf, _method
                              in non_english_rows},
            # Which rows were looked at, so a consumer can tell "English" from
            # "too short to judge" without re-deriving the length filter.
            'analyzed_mask': df.index.isin(text_series.index),
            **run_info,
        }
    
    def check_columns(self, df: pd.DataFrame, columns: List[str], 
                      min_length: int = 15, min_confidence: float = 0.85,
                      mode: str = 'all') -> List[Dict[str, Any]]:
        """Check multiple columns for non-English text."""
        return [self.check_column(df, col, min_length, min_confidence, mode) for col in columns]
    
    def detect_language_series(self, series: pd.Series,
                               min_confidence: float = 0.75,
                               mode: str = 'all') -> pd.Series:
        """
        Detect language for every row of one column, aligned to its index.

        Detects over `unique()` and maps back, which matters more here than
        anywhere else in the app: `_detection_cache` is capped at 10,000
        entries, so a row-by-row pass over a column with more distinct texts
        than that stops getting any cache hits at all and pays the full model
        cost per row. Deduplicating first bounds the work by the number of
        distinct texts instead - the same trick as LanguageChecker.clean_series.

        Blank rows come back as '' rather than 'English': nothing was judged.
        """
        text = series.astype(str).fillna('')
        stripped = text.str.strip()
        populated = stripped != ''

        answers: Dict[str, str] = {}
        for value in stripped[populated].unique():
            answers[value] = self.detect_language_for_row(
                value, min_confidence, mode)

        return stripped.map(answers).fillna('')

    def get_language_columns(self, df: pd.DataFrame, columns: List[str],
                             min_confidence: float = 0.75,
                             mode: str = 'all') -> pd.DataFrame:
        """
        Detect language for each row in specified columns.
        Returns a DataFrame with language detection columns that can be joined to the original data.

        Args:
            df: The DataFrame to analyze
            columns: List of column names to check for languages
            min_confidence: Minimum confidence threshold for Lingua detection
            mode: 'all', 'lingua_only', or 'rules_only'

        Returns:
            DataFrame with columns: original index + '{col}_Language' for each column
        """
        result_data = {'_original_index': df.index.tolist()}

        for col in columns:
            if col not in df.columns:
                continue
            result_data[f'{col}_Language'] = self.detect_language_series(
                df[col], min_confidence, mode).tolist()

        return pd.DataFrame(result_data)

    def add_language_columns_to_df(self, df: pd.DataFrame, columns: List[str],
                                    min_confidence: float = 0.75,
                                    mode: str = 'all') -> pd.DataFrame:
        """
        Add language detection columns to a DataFrame copy.
        For each specified column, adds a new column '{column}_Language' with detected language.

        Args:
            df: The original DataFrame
            columns: List of columns to detect language in
            min_confidence: Minimum confidence for Lingua detection
            mode: 'all', 'lingua_only', or 'rules_only'

        Returns:
            A new DataFrame with added language columns
        """
        df_copy = df.copy()

        for col in columns:
            if col not in df_copy.columns:
                continue
            df_copy[f'{col}_Language'] = self.detect_language_series(
                df_copy[col], min_confidence, mode)

        return df_copy

    def format_result(self, result: Dict[str, Any]) -> List[str]:
        """Format a column check result as human-readable strings."""
        lines = []
        col = result['column']
        
        if 'error' in result:
            lines.append(f"⚠️ '{col}': {result['error']}")
            return lines
        
        if 'warning' in result:
            lines.append(f"⚠️ '{col}': {result['warning']}")
            return lines
        
        non_english = result['non_english_count']
        pct = result['percentage']
        total = result['total_analyzed']
        original = result['total_original']
        
        if non_english > 0:
            # Format language counts
            lang_summary = ", ".join([
                f"{self.LANG_NAMES.get(l, l)}:{c}" 
                for l, c in sorted(result['languages'].items(), key=lambda x: -x[1])[:6]
            ])
            
            lines.append(f"❌ '{col}': {non_english} rows ({pct:.1f}%) contain non-English text.")
            lines.append(f"   Analyzed: {total} of {original} rows")
            lines.append(f"   Languages: {lang_summary}")
            
            # Show samples
            for i, (idx, text, lang, conf, method) in enumerate(result['samples'][:5], 1):
                truncated = text[:80] + "..." if len(text) > 80 else text
                lang_name = self.LANG_NAMES.get(lang, lang)
                lines.append(f"   Sample {i} [{lang_name}, via {method}]: {truncated}")
        else:
            lines.append(f"✅ '{col}': All {total} entries appear to be in English.")
            lines.append(f"   Analyzed: {total} of {original} rows")
        
        return lines
