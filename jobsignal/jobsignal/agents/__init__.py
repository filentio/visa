from .collector import CollectorAgent
from .dedup import DeduplicatorAgent
from .matcher import MatcherAgent
from .parser import Parser as _Parser
from .raw_dedup import RawDeduplicatorAgent
from jobsignal.config import AppConfig

class ParserAgent(_Parser):
    """Wrapper to make Parser compatible with BaseAgent interface."""
    def __init__(self, config: AppConfig) -> None:
        super().__init__()
        self.config = config
    @property
    def name(self):
        return "parser"

__all__ = ["CollectorAgent", "RawDeduplicatorAgent", "ParserAgent",
           "DeduplicatorAgent", "MatcherAgent"]
