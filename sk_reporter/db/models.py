"""ORM-модели планирования."""

from __future__ import annotations

from datetime import datetime

from sqlalchemy import DateTime, String, Text, func
from sqlalchemy.orm import Mapped, mapped_column

from sk_reporter.db.base import Base


class Personnel(Base):
    __tablename__ = "personnel"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    fio: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    phone: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    position: Mapped[str] = mapped_column(Text, nullable=False, default="")
    control_mode: Mapped[str] = mapped_column(Text, nullable=False, default="")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )
