"""ORM-модели планирования."""

from __future__ import annotations

from datetime import datetime

from sqlalchemy import Boolean, DateTime, ForeignKey, String, Text, func
from sqlalchemy.dialects.postgresql import JSONB
from sqlalchemy.orm import Mapped, mapped_column

from sk_reporter.db.base import Base


class Contractor(Base):
    __tablename__ = "contractors"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    name: Mapped[str] = mapped_column(String(255), nullable=False, index=True)
    template_stem: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    file_label: Mapped[str] = mapped_column(String(128), nullable=False, default="", index=True)
    inspection_type: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    gen_contractor: Mapped[str] = mapped_column(String(512), nullable=False, default="")
    sub_contractor: Mapped[str] = mapped_column(String(512), nullable=False, default="")
    contract_no: Mapped[str] = mapped_column(String(128), nullable=False, default="")
    contact_person: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    contact_phone: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    contact_fax: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    contact_email: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    extra_info: Mapped[str] = mapped_column(Text, nullable=False, default="")
    note_discrepancy: Mapped[str] = mapped_column(Text, nullable=False, default="")
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )


class Project(Base):
    __tablename__ = "projects"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    contractor_id: Mapped[str | None] = mapped_column(
        String(64), ForeignKey("contractors.id", ondelete="SET NULL"), nullable=True, index=True
    )
    title: Mapped[str] = mapped_column(String(512), nullable=False, default="")
    object_name: Mapped[str] = mapped_column(Text, nullable=False, default="")
    vor_file: Mapped[str] = mapped_column(String(512), nullable=False, default="")
    tl_file: Mapped[str] = mapped_column(String(512), nullable=False, default="")
    content: Mapped[dict | None] = mapped_column(JSONB, nullable=True)
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )


class ProjectEngineer(Base):
    __tablename__ = "project_engineers"

    project_id: Mapped[str] = mapped_column(
        String(64), ForeignKey("projects.id", ondelete="CASCADE"), primary_key=True
    )
    person_id: Mapped[str] = mapped_column(
        String(64), ForeignKey("personnel.id", ondelete="CASCADE"), primary_key=True
    )
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())


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


class OtkkCard(Base):
    __tablename__ = "otkk_cards"

    id: Mapped[str] = mapped_column(String(64), primary_key=True)
    file_name: Mapped[str] = mapped_column(String(512), nullable=False)
    code: Mapped[str] = mapped_column(String(64), nullable=False, default="")
    title: Mapped[str] = mapped_column(Text, nullable=False, default="")
    content: Mapped[dict | None] = mapped_column(JSONB, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), server_default=func.now())
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )
