from __future__ import annotations

from datetime import UTC, datetime
from typing import ClassVar

from sqlalchemy import DateTime, Text
from sqlalchemy.dialects.postgresql import ARRAY, BIGINT
from sqlalchemy.orm import Mapped, mapped_column

from config import settings
from sql.database import Base


class SolutionsDoc(Base):
    __tablename__ = settings.solution_docs_table_name
    __table_args__: ClassVar[dict[str]] = {
        "schema": settings.solution_docs_table_schema
    }

    id: Mapped[int] = mapped_column(
        BIGINT, primary_key=True, index=True, nullable=False
    )
    sp_id: Mapped[str] = mapped_column(Text, unique=True, nullable=False)
    file_type: Mapped[str] = mapped_column(Text)
    name: Mapped[str] = mapped_column(Text, nullable=False, index=True)
    content_text: Mapped[str | None] = mapped_column(Text)
    img_metadata: Mapped[list[str] | None] = mapped_column(
        ARRAY(Text, dimensions=1)
    )
    keywords: Mapped[list[str] | None] = mapped_column(
        ARRAY(Text, dimensions=1), index=True
    )
    summary: Mapped[str | None] = mapped_column(Text)
    file_path: Mapped[str] = mapped_column(Text, index=True)
    web_url: Mapped[str] = mapped_column(Text)
    created_datetime: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False
    )
    last_modified_datetime: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False
    )
    last_synced: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), default=lambda: datetime.now(UTC)
    )
