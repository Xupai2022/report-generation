from __future__ import annotations

import importlib
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import patch


class FakeAssetCollection:
    def __init__(self, docs: list[dict]):
        self.docs = docs
        self.find_calls: list[tuple[dict, dict, int]] = []

    def find(self, query: dict, projection: dict, batch_size: int | None = None):
        self.find_calls.append((query, projection, batch_size or 0))
        target_assets = set(query["asset"]["$in"])
        return [doc for doc in self.docs if doc.get("asset") in target_assets]


class FakeDatabase:
    def __init__(self, asset_collection: FakeAssetCollection):
        self.asset_collection = asset_collection

    def __getitem__(self, name: str):
        if name != "Assets":
            raise KeyError(name)
        return self.asset_collection


class EventAssetMappingTestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        fake_pymongo = SimpleNamespace(MongoClient=object)
        with patch.dict(sys.modules, {"pymongo": fake_pymongo}):
            cls.collector_module = importlib.import_module(
                "mss_ai_ppt_sample_assets.backend.services.data_ingestion.collectors.mongo_collector"
            )

    def test_attach_event_asset_security_domains_maps_single_ip(self) -> None:
        asset_collection = FakeAssetCollection(
            [
                {"asset": "10.0.0.1", "security_domain": "内网"},
                {"asset": "10.0.0.2", "security_domain": "外网"},
            ]
        )
        database = FakeDatabase(asset_collection)
        event_docs = [{"host_ip": "10.0.0.1"}, {"host_ip": "10.0.0.3"}]

        self.collector_module._attach_event_asset_security_domains(
            event_docs,
            database=database,
            company_id="tenant-1",
            asset_collection="Assets",
            batch_size=1000,
        )

        self.assertEqual(event_docs[0]["内网外网资产"], "内网")
        self.assertEqual(event_docs[1]["内网外网资产"], "")
        self.assertEqual(
            asset_collection.find_calls[0][0],
            {
                "company_id": "tenant-1",
                "is_deleted": 0,
                "asset": {"$in": ["10.0.0.1", "10.0.0.3"]},
            },
        )

    def test_attach_event_asset_security_domains_maps_multiple_ips(self) -> None:
        asset_collection = FakeAssetCollection(
            [
                {"asset": "10.0.0.1", "security_domain": "内网"},
                {"asset": "10.0.0.2", "security_domain": "DMZ"},
            ]
        )
        database = FakeDatabase(asset_collection)
        event_docs = [{"host_ip": "10.0.0.1,10.0.0.2"}]

        self.collector_module._attach_event_asset_security_domains(
            event_docs,
            database=database,
            company_id="tenant-1",
            asset_collection="Assets",
            batch_size=1000,
        )

        self.assertEqual(event_docs[0]["内网外网资产"], "内网、DMZ")

    def test_attach_event_asset_security_domains_deduplicates_repeated_ips(self) -> None:
        asset_collection = FakeAssetCollection(
            [{"asset": "10.0.0.1", "security_domain": "内网"}]
        )
        database = FakeDatabase(asset_collection)
        event_docs = [{"host_ip": "10.0.0.1, 10.0.0.1"}]

        self.collector_module._attach_event_asset_security_domains(
            event_docs,
            database=database,
            company_id="tenant-1",
            asset_collection="Assets",
            batch_size=1000,
        )

        self.assertEqual(event_docs[0]["内网外网资产"], "内网")


if __name__ == "__main__":
    unittest.main()
