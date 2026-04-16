from __future__ import annotations

from datetime import datetime
import logging
from typing import Any, Dict

from pymongo import MongoClient

from .... import config
from ..models import MongoCollectionConfig, MongoIngestionRequest

logger = logging.getLogger(__name__)

ALARM_EVENT_STATUS_DISPLAY: Dict[str, str] = {
    "inited": "未处置",
    "finished": "已完成",
    "reject_ignore": "已忽略（驳回）",
    "rejected": "已驳回",
    "disposal": "处置中",
    "link_event": "已关联告警",
    "relate_event": "生成事件",
    "ignore": "已忽略",
    "white": "已加白",
    "auto_event": "生成事件（自动）",
    "auto_ignore": "已忽略（自动）",
}

ATTACK_STATE_DISPLAY: Dict[int, str] = {
    0: "judge（待研判）",
    1: "fail（失败）",
    2: "succ（成功）",
    3: "compromised（失陷）",
    4: "anomaly（异常）",
    5: "attempt（尝试）",
}

ATTACK_DIRECTION_DISPLAY: Dict[int, str] = {
    0: "未知",
    1: "内-外",
    2: "外-内",
    3: "内-内",
    4: "外-外",
}

ALARM_SERVICE_STATUS_DISPLAY: Dict[int, str] = {
    7: "服务内（7*24H）",
    8: "服务外",
    9: "服务内（5*8H）",
    10: "全部服务",
}

PUSH_STATUS_DISPLAY: Dict[int, str] = {
    1: "已通告",
    -1: "未通告",
}

EVENT_EVENT_STATUS_DISPLAY: Dict[str, str] = {
    "inited": "未处置",
    "disposal": "处置中",
    "suspend": "已暂停",
    "finished": "已完成",
    "accept_risk": "接受风险",
    "protected": "已防护",
    "announced": "已通告",
    "rejected": "已驳回",
}

EVENT_SERVICE_STATUS_DISPLAY: Dict[int, str] = {
    0: "服务内（7*24H）",
    1: "服务外",
    3: "服务内（5*8H）",
    -1: "全部服务",
}

ALARM_PROJECTION: Dict[str, int] = {
    "alarm_name": 1,
    "asset": 1,
    "manage_type_name": 1,
    "manage_sub_type_name": 1,
    "event_status": 1,
    "first_time": 1,
    "latest_time": 1,
    "service_status": 1,
    "create_time": 1,
    "attack_state": 1,
    "attack_direction": 1,
    "current_operate_time": 1,
    "rejected_event_id": 1,
    "current_operator": 1,
    "reject_reason": 1,
    "_id": 0,
}

EVENT_PROJECTION: Dict[str, int] = {
    "create_time": 1,
    "manage_type": 1,
    "manage_sub_type": 1,
    "host_ip": 1,
    "event_status": 1,
    "service_status": 1,
    "latest_time": 1,
    "checkout_time": 1,
    "dispose_time": 1,
    "contain_time": 1,
    "finished_time": 1,
    "incidence": 1,
    "update_protected_time": 1,
    "update_announced_time": 1,
    "update_accept_risk_time": 1,
    "push_status": 1,
    "wechat_push_time": 1,
    "_id": 0,
}


class SOARMongoCollector:
    """Collect raw alarm/event documents from the SOAR MongoDB source."""

    def __init__(
        self,
        mongo_uri: str | None = None,
        collections: MongoCollectionConfig | None = None,
        connect_timeout_ms: int | None = None,
    ) -> None:
        self.mongo_uri = (mongo_uri or config.settings.soar_mongo_uri or "").strip()
        if not self.mongo_uri:
            raise ValueError(
                "SOAR MongoDB connection is not configured. "
                "Set SOAR_MONGO_URI or the split SOAR_MONGO_* environment variables."
            )
        self.collections = collections or MongoCollectionConfig(
            database_name=config.settings.soar_mongo_database,
            alarm_collection=config.settings.soar_mongo_alarm_collection,
            event_collection=config.settings.soar_mongo_event_collection,
        )
        self.connect_timeout_ms = int(
            connect_timeout_ms or config.settings.soar_mongo_connect_timeout_ms
        )

    def collect(self, request: MongoIngestionRequest) -> Dict[str, Any]:
        query = {
            "company_id": request.company_id,
            "create_time": {
                "$gte": request.start_time,
                "$lte": request.end_time,
            },
        }

        logger.info(
            "Collecting SOAR Mongo data | db=%s company_id=%s start=%s end=%s",
            self.collections.database_name,
            request.company_id,
            request.start_time.isoformat(),
            request.end_time.isoformat(),
        )

        client = MongoClient(
            self.mongo_uri,
            serverSelectionTimeoutMS=self.connect_timeout_ms,
            connectTimeoutMS=self.connect_timeout_ms,
        )
        try:
            database = client[self.collections.database_name]
            alarm_docs = [
                _transform_alarm_doc(doc)
                for doc in database[self.collections.alarm_collection].find(
                    query, ALARM_PROJECTION
                )
            ]
            event_docs = list(
                _transform_event_doc(doc)
                for doc in database[self.collections.event_collection].find(
                    query, EVENT_PROJECTION
                )
            )
        finally:
            client.close()

        return {
            "meta": {
                "company_id": request.company_id,
                "start_time": _serialize_datetime(request.start_time),
                "end_time": _serialize_datetime(request.end_time),
                "database_name": self.collections.database_name,
                "alarm_collection": self.collections.alarm_collection,
                "event_collection": self.collections.event_collection,
                "alarm_count": len(alarm_docs),
                "event_count": len(event_docs),
                "extracted_at": _serialize_datetime(datetime.utcnow()),
            },
            "alarm": alarm_docs,
            "event": event_docs,
        }


def _serialize_datetime(value: datetime) -> str:
    if value.tzinfo is None:
        return value.isoformat() + "Z"
    return value.isoformat()


def _transform_alarm_doc(doc: Dict[str, Any]) -> Dict[str, Any]:
    transformed = dict(doc)
    transformed["event_status"] = ALARM_EVENT_STATUS_DISPLAY.get(
        transformed.get("event_status"),
        transformed.get("event_status"),
    )
    transformed["attack_state"] = ATTACK_STATE_DISPLAY.get(
        transformed.get("attack_state"),
        transformed.get("attack_state"),
    )
    transformed["attack_direction"] = ATTACK_DIRECTION_DISPLAY.get(
        transformed.get("attack_direction"),
        transformed.get("attack_direction"),
    )
    transformed["service_status"] = ALARM_SERVICE_STATUS_DISPLAY.get(
        transformed.get("service_status"),
        transformed.get("service_status"),
    )

    if "reject_reason" not in transformed or transformed.get("reject_reason") in (None, ""):
        transformed["reject_reason"] = "占位"

    return transformed


def _transform_event_doc(doc: Dict[str, Any]) -> Dict[str, Any]:
    transformed = dict(doc)
    transformed["event_status"] = EVENT_EVENT_STATUS_DISPLAY.get(
        transformed.get("event_status"),
        transformed.get("event_status"),
    )
    transformed["service_status"] = EVENT_SERVICE_STATUS_DISPLAY.get(
        transformed.get("service_status"),
        transformed.get("service_status"),
    )
    transformed["push_status"] = PUSH_STATUS_DISPLAY.get(
        transformed.get("push_status"),
        transformed.get("push_status"),
    )
    return transformed
