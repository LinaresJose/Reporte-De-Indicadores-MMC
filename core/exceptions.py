"""Custom exceptions for the LYNX reporting system."""


class LynxBaseError(Exception):
    """Base exception for all LYNX-specific errors.

    All custom exceptions in this module inherit from this class, allowing
    callers to catch any LYNX error with a single ``except LynxBaseError``.

    Attributes:
        detail: An optional human-readable string with additional context
            about the error condition.
    """

    def __init__(self, message: str = "", detail: str = "") -> None:
        """Initialise LynxBaseError.

        Args:
            message: Primary error message passed to :class:`Exception`.
            detail: Optional supplementary information about the error.
        """
        super().__init__(message)
        self.detail: str = detail

    def __repr__(self) -> str:
        """Return a detailed developer-facing representation."""
        cls = type(self).__name__
        base = str(self)
        if self.detail:
            return f"{cls}({base!r}, detail={self.detail!r})"
        return f"{cls}({base!r})"


class ETLError(LynxBaseError):
    """Raised when a source-file read or transformation operation fails.

    This is the base class for all errors that occur during the
    Extract-Transform-Load phase of the pipeline, such as missing files,
    unreadable formats, or unexpected schemas.
    """


class EngineError(LynxBaseError):
    """Raised during KPI calculation in the reporting engine.

    Use this exception when a business-logic computation encounters an
    unrecoverable condition, such as referencing data that was expected
    to be present after the ETL phase.
    """


class ReportError(LynxBaseError):
    """Raised during report generation (Excel, PDF, or CSV output).

    Covers errors such as file-system permission failures, missing
    templates, or unexpected rendering states.
    """


class MaintenanceError(LynxBaseError):
    """Raised during maintenance operations.

    Covers errors arising from monthly data loaders for budgets and
    collections when configuration files are malformed or unavailable.
    """


class MissingColumnError(ETLError):
    """Raised when a required column is absent from a source DataFrame.

    Attributes:
        column_name: The exact column name that could not be found.
        detail: Additional context, e.g. the file name that was being parsed.
    """

    def __init__(self, column_name: str, detail: str = "") -> None:
        """Initialise MissingColumnError.

        Args:
            column_name: The name of the column that is missing.
            detail: Optional supplementary information (e.g. source file path).
        """
        message = f"Required column not found: '{column_name}'"
        super().__init__(message, detail=detail)
        self.column_name: str = column_name

    def __repr__(self) -> str:
        """Return a detailed developer-facing representation."""
        cls = type(self).__name__
        return (
            f"{cls}(column_name={self.column_name!r}, detail={self.detail!r})"
        )


class EmptySourceError(ETLError):
    """Raised when a source file contains no data rows after loading.

    This signals that the file was readable and correctly structured, but
    the resulting DataFrame is empty — which would produce a meaningless report.
    """


class VendorNotFoundError(EngineError):
    """Raised when a vendor present in source data is absent from Zonas.

    Every vendor referenced in the sales/collections files must have a
    corresponding entry in the Zonas configuration table.  If one is
    missing, KPI calculations cannot be completed for that vendor.
    """
