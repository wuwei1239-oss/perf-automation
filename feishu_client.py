# -*- coding: utf-8 -*-
"""
飞书多维表格 API 客户端
- 自动管理 tenant_access_token
- list_records / create_records / update_records / delete_records
- 批量操作支持
"""
import os, time, requests, json

APP_ID = os.environ.get("FEISHU_APP_ID", "cli_a97bd5c8ec795bc6")
APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "Z0U6Bqc04l7zYoKJlEMoPfH60jePIica")
BASE_ID = os.environ.get("FEISHU_BASE_ID", "EtxBbORRna1Sdms4gxpcH6SEnNh")
BASE_URL = "https://open.feishu.cn/open-apis/bitable/v1"

# 表名 -> table_id 映射
TABLES = {
    "系数规则": "tbl4Mh9i1fbpYmmt",
    "车间人员": "tbl1AdIv8VJAGkwo",
    "品种工序规则": "tbl17rYHOsFHuq8D",
    "每日报工流水账": "tblkDy41lNZC2Z9B",
    "个人绩效明细表": "tblSSKwtVTK4fCom",
    "人员工序绑定表": "tblrqjKQTkE3XEID",
    "绩效考核表": "tbl8Qg5pvMYQkO5V",
}


class FeishuClient:
    def __init__(self, app_id=None, app_secret=None, base_id=None):
        self.app_id = app_id or APP_ID
        self.app_secret = app_secret or APP_SECRET
        self.base_id = base_id or BASE_ID
        self._token = None
        self._token_expires_at = 0

    def _ensure_token(self):
        if self._token and time.time() < self._token_expires_at - 60:
            return
        resp = requests.post(
            "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
            json={"app_id": self.app_id, "app_secret": self.app_secret},
            timeout=15,
        )
        data = resp.json()
        if resp.status_code != 200 or data.get("code", -1) != 0:
            raise RuntimeError(f"获取token失败: {data}")
        self._token = data["tenant_access_token"]
        self._token_expires_at = time.time() + data.get("expire", 7200)

    @property
    def headers(self):
        self._ensure_token()
        return {
            "Authorization": f"Bearer {self._token}",
            "Content-Type": "application/json; charset=utf-8",
        }

    # ------------------------------------------------------------------
    # 记录操作
    # ------------------------------------------------------------------
    def list_records(self, table_id, page_size=500, page_token=None, view_id=None, **kwargs):
        """列出记录（自动翻页返回全部）"""
        all_items = []
        params = {"page_size": min(page_size, 500)}
        if page_token:
            params["page_token"] = page_token
        if view_id:
            params["view_id"] = view_id

        while True:
            r = requests.get(
                f"{BASE_URL}/apps/{self.base_id}/tables/{table_id}/records",
                headers=self.headers,
                params=params,
                timeout=30,
            )
            j = r.json()
            if r.status_code != 200 or j.get("code", -1) != 0:
                raise RuntimeError(f"list_records失败: HTTP{r.status_code} {j}")
            data = j.get("data", {})
            items = data.get("items", [])
            all_items.extend(items)
            if not data.get("has_more"):
                break
            params["page_token"] = data.get("page_token", "")
        return all_items

    def create_records(self, table_id, records):
        """批量创建记录，每批最多500条。records = [{"fields": {...}}, ...]"""
        BATCH = 500
        created = []
        for i in range(0, len(records), BATCH):
            batch = records[i:i + BATCH]
            r = requests.post(
                f"{BASE_URL}/apps/{self.base_id}/tables/{table_id}/records/batch_create",
                headers=self.headers,
                json={"records": batch},
                timeout=60,
            )
            j = r.json()
            if r.status_code != 200 or j.get("code", -1) != 0:
                raise RuntimeError(f"create_records失败: HTTP{r.status_code} {j}")
            created.extend(j.get("data", {}).get("records", []))
        return created

    def update_records(self, table_id, records):
        """批量更新记录，records = [{"record_id": "xxx", "fields": {...}}, ...]"""
        BATCH = 500
        updated = []
        for i in range(0, len(records), BATCH):
            batch = records[i:i + BATCH]
            r = requests.post(
                f"{BASE_URL}/apps/{self.base_id}/tables/{table_id}/records/batch_update",
                headers=self.headers,
                json={"records": batch},
                timeout=60,
            )
            j = r.json()
            if r.status_code != 200 or j.get("code", -1) != 0:
                raise RuntimeError(f"update_records失败: HTTP{r.status_code} {j}")
            updated.extend(j.get("data", {}).get("records", []))
        return updated

    def delete_records(self, table_id, record_ids):
        """批量删除记录"""
        BATCH = 500
        for i in range(0, len(record_ids), BATCH):
            r = requests.post(
                f"{BASE_URL}/apps/{self.base_id}/tables/{table_id}/records/batch_delete",
                headers=self.headers,
                json={"records": record_ids[i:i + BATCH]},
                timeout=60,
            )
            j = r.json()
            if r.status_code != 200 or j.get("code", -1) != 0:
                raise RuntimeError(f"delete_records失败: HTTP{r.status_code} {j}")

    def get_all_records(self, table_name):
        """根据中文表名获取全部记录"""
        tid = TABLES.get(table_name)
        if not tid:
            raise ValueError(f"未知表名: {table_name}")
        return self.list_records(tid)

    # ------------------------------------------------------------------
    # 字段操作
    # ------------------------------------------------------------------
    def list_fields(self, table_id):
        r = requests.get(
            f"{BASE_URL}/apps/{self.base_id}/tables/{table_id}/fields",
            headers=self.headers,
            timeout=15,
        )
        j = r.json()
        if r.status_code != 200 or j.get("code", -1) != 0:
            raise RuntimeError(f"list_fields失败: HTTP{r.status_code} {j}")
        return j.get("data", {}).get("items", [])


# ============================================================
# 便捷工具函数
# ============================================================
def build_coeff_map(client: FeishuClient = None):
    """从系数规则表构建 岗位身份->系数 映射"""
    if client is None:
        client = FeishuClient()
    records = client.get_all_records("系数规则")
    mapping = {}
    for r in records:
        f = r.get("fields", {})
        role = f.get("岗位身份", "")
        coeff_str = f.get("系数", "1.0")
        try:
            coeff = float(coeff_str)
        except (ValueError, TypeError):
            coeff = 1.0
        if role:
            mapping[role] = coeff
    return mapping


def build_person_coeff_map(client: FeishuClient = None):
    """从人员工序绑定表构建 姓名->系数 映射"""
    if client is None:
        client = FeishuClient()
    records = client.get_all_records("人员工序绑定表")
    mapping = {}
    for r in records:
        f = r.get("fields", {})
        name = f.get("姓名", "")
        if not name:
            continue
        coeff_raw = f.get("系数")
        if isinstance(coeff_raw, list) and coeff_raw:
            coeff_raw = coeff_raw[0]
        try:
            coeff = float(coeff_raw) if coeff_raw else 1.0
        except (ValueError, TypeError):
            coeff = 1.0
        mapping[name] = coeff
    return mapping


def build_rules_map(client: FeishuClient = None):
    """从品种工序规则表构建 专属号->规则行 映射（对齐perf_engine语义）"""
    if client is None:
        client = FeishuClient()
    records = client.get_all_records("品种工序规则")
    rules = {}
    for r in records:
        f = r.get("fields", {})
        code = f.get("专属号（批号前缀）", "")
        if not code:
            continue
        rules[code] = f
    return rules


def build_person_list(client: FeishuClient = None):
    """从绩效考核表获取人员列表及现有月累计"""
    if client is None:
        client = FeishuClient()
    records = client.get_all_records("绩效考核表")
    persons = []
    for r in records:
        f = r.get("fields", {})
        name = f.get("姓名", "")
        if name:
            persons.append({
                "record_id": r.get("record_id", ""),
                "name": name,
                "coeff": float(f.get("系数", 1) or 1),
                "月累计": float(f.get("月累计", 0) or 0),
                "生产得分": float(f.get("生产得分", 0) or 0),
                "加减分": float(f.get("加减分", 0) or 0),
            })
    return persons


if __name__ == "__main__":
    c = FeishuClient()
    print(f"Token: {c._token[:20]}..." if c._token else "Token: 未获取")
    c._ensure_token()
    print(f"Token: {c._token[:20]}...")

    # 测试读取品种规则
    rules = build_rules_map(c)
    print(f"\n品种规则: {len(rules)} 条")
    for code in sorted(rules)[:5]:
        print(f"  {code}: {rules[code].get('产品名称', '')}")

    # 测试读取系数
    coeff_map = build_coeff_map(c)
    print(f"\n系数规则: {len(coeff_map)} 条")
    for role, coeff in sorted(coeff_map.items())[:5]:
        print(f"  {role}: {coeff}")

    # 测试人员
    persons = build_person_list(c)
    print(f"\n绩效考核表人员: {len(persons)} 人")
    for p in persons[:5]:
        print(f"  {p['name']} 系数={p['coeff']} 月累计={p['月累计']}")