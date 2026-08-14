class SOPReporterError(Exception):
    """Base exception for expected application failures."""


class ConfigurationError(SOPReporterError):
    """Raised when a YAML configuration value is invalid."""


class CredentialsError(SOPReporterError):
    """Raised when Gmail credentials cannot be obtained."""


class CredentialsCancelledError(CredentialsError):
    """Raised when the first-run credential dialog is cancelled."""


class EmailClientError(SOPReporterError):
    """Raised for IMAP or MIME-processing errors."""


class ExtractionError(SOPReporterError):
    """Raised when a workbook cannot be transformed by the configured rules."""


class ReportBuildError(SOPReporterError):
    """Raised when the output workbook cannot be built."""


class PrintError(SOPReporterError):
    """Raised when Excel COM cannot print the generated report."""

    def __init__(self, message: str, *, print_invoked: bool = False) -> None:
        super().__init__(message)
        self.print_invoked = print_invoked


class StateStoreError(SOPReporterError):
    """Raised when deduplication state cannot be read or persisted."""


class UpdateError(SOPReporterError):
    """Raised when a desktop self-update cannot be checked, downloaded, or applied."""
