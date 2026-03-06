import React, { useEffect, useMemo, useState } from 'react';
import { BarChart3, LogOut, RefreshCw, Search, Trash2 } from 'lucide-react';
import { AdminApi } from '../api/main';
import type { AdminJob, AdminJobDetailsResp, AdminStatisticsResp } from '../types/api';

function fmtDate(input?: string) {
  if (!input) return '--';
  try {
    return new Date(input).toLocaleString('zh-CN');
  } catch {
    return input;
  }
}

export function AdminApp() {
  const [jobs, setJobs] = useState<AdminJob[]>([]);
  const [stats, setStats] = useState<AdminStatisticsResp | null>(null);
  const [details, setDetails] = useState<AdminJobDetailsResp | null>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const [search, setSearch] = useState('');
  const [rating, setRating] = useState('');
  const [status, setStatus] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());

  const selectedJobId = useMemo(() => details?.job?.job_id || '', [details]);

  const ensureAuth = async () => {
    try {
      const result = await AdminApi.verify();
      if (!result.authenticated) {
        window.location.href = '/ui/login.html';
      }
    } catch {
      window.location.href = '/ui/login.html';
    }
  };

  const loadStats = async () => {
    try {
      const result = await AdminApi.stats();
      setStats(result);
    } catch (err) {
      console.error(err);
    }
  };

  const loadJobs = async () => {
    setLoading(true);
    setError('');
    try {
      const params = new URLSearchParams({
        page: '1',
        limit: '50',
      });
      if (search.trim()) params.set('search', search.trim());
      if (rating) params.set('rating', rating);
      if (status) params.set('status', status);
      if (dateFrom) params.set('date_from', `${dateFrom}T00:00:00Z`);
      if (dateTo) params.set('date_to', `${dateTo}T23:59:59Z`);

      const result = await AdminApi.listJobs(params);
      setJobs(result.jobs || []);
      setSelectedIds(new Set());
    } catch (err) {
      setError(err instanceof Error ? err.message : '加载任务失败');
    } finally {
      setLoading(false);
    }
  };

  const loadDetails = async (jobId: string) => {
    try {
      const result = await AdminApi.getJobDetail(jobId);
      setDetails(result);
    } catch (err) {
      setError(err instanceof Error ? err.message : '加载详情失败');
    }
  };

  const logout = async () => {
    try {
      await AdminApi.logout();
    } finally {
      window.location.href = '/ui/login.html';
    }
  };

  const toggleSelected = (jobId: string) => {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(jobId)) next.delete(jobId);
      else next.add(jobId);
      return next;
    });
  };

  const deleteSelected = async () => {
    if (selectedIds.size === 0) return;
    const ok = window.confirm(`确认删除 ${selectedIds.size} 个任务？`);
    if (!ok) return;

    setLoading(true);
    try {
      await AdminApi.deleteJobs(Array.from(selectedIds));
      await Promise.all([loadJobs(), loadStats()]);
      setDetails(null);
    } catch (err) {
      setError(err instanceof Error ? err.message : '删除失败');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void ensureAuth().then(async () => {
      await Promise.all([loadJobs(), loadStats()]);
    });
  }, []);

  useEffect(() => {
    const timer = window.setTimeout(() => {
      void loadJobs();
    }, 400);
    return () => window.clearTimeout(timer);
  }, [search, rating, status, dateFrom, dateTo]);

  return (
    <div className="admin-page">
      <header className="admin-header">
        <div>
          <h1>MSS AI PPT 管理后台</h1>
          <p>任务监控、反馈统计与批量管理</p>
        </div>
        <button className="btn-secondary" onClick={logout}><LogOut size={14} /> 退出登录</button>
      </header>

      <section className="admin-toolbar">
        <label className="search-input">
          <Search size={14} />
          <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="搜索任务ID、模板、输入源" />
        </label>
        <select value={rating} onChange={(e) => setRating(e.target.value)}>
          <option value="">评分：全部</option>
          <option value="liked">点赞</option>
          <option value="disliked">点踩</option>
          <option value="unrated">未评分</option>
        </select>
        <select value={status} onChange={(e) => setStatus(e.target.value)}>
          <option value="">状态：全部</option>
          <option value="completed">已完成</option>
          <option value="failed">失败</option>
          <option value="running">运行中</option>
          <option value="pending">等待中</option>
        </select>
        <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} />
        <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} />
        <button className="btn-secondary" onClick={() => void loadJobs()} disabled={loading}><RefreshCw size={14} /> 刷新</button>
        <button className="btn-danger" onClick={deleteSelected} disabled={loading || selectedIds.size === 0}><Trash2 size={14} /> 删除选中</button>
      </section>

      {error && <div className="admin-error">{error}</div>}

      <div className="admin-body">
        <aside className="admin-list">
          {jobs.length === 0 && <div className="empty">暂无任务</div>}
          {jobs.map((job) => {
            const selected = selectedJobId === job.job_id;
            return (
              <div key={job.job_id} className={`job-item ${selected ? 'selected' : ''}`} onClick={() => void loadDetails(job.job_id)}>
                <input
                  type="checkbox"
                  checked={selectedIds.has(job.job_id)}
                  onChange={() => toggleSelected(job.job_id)}
                  onClick={(e) => e.stopPropagation()}
                />
                <div className="job-main">
                  <div className="job-id">{job.job_id}</div>
                  <div className="job-meta">
                    <span className={`pill status-${job.status}`}>{job.status}</span>
                    <span>{job.rating?.rating === 'liked' ? '👍' : job.rating?.rating === 'disliked' ? '👎' : '—'}</span>
                  </div>
                  <div className="job-time">{fmtDate(job.created_at)}</div>
                </div>
              </div>
            );
          })}
        </aside>

        <main className="admin-detail">
          {!details ? (
            <div className="empty">请选择左侧任务查看详情</div>
          ) : (
            <>
              <section className="panel">
                <h3>任务详情</h3>
                <div className="kv-grid">
                  <div><span>任务ID</span><strong>{details.job.job_id}</strong></div>
                  <div><span>模板</span><strong>{details.job.template_id}</strong></div>
                  <div><span>输入源</span><strong>{details.job.input_id}</strong></div>
                  <div><span>状态</span><strong>{details.job.status}</strong></div>
                  <div><span>创建时间</span><strong>{fmtDate(details.job.created_at)}</strong></div>
                  <div><span>AI模型</span><strong>{(details.job.metadata?.ai_model as string) || '--'}</strong></div>
                </div>
              </section>

              <section className="panel">
                <h3>预览</h3>
                <div className="preview-stack">
                  {(details.preview_urls || []).map((url, idx) => (
                    <img key={url + idx} src={url} alt={`slide-${idx + 1}`} />
                  ))}
                  {(details.preview_urls || []).length === 0 && <div className="empty">无预览图片</div>}
                </div>
              </section>

              <section className="panel">
                <h3>用户反馈</h3>
                {!details.job.rating?.rating ? (
                  <div className="empty">暂无评分</div>
                ) : (
                  <div className="rating-show">
                    <div>{details.job.rating.rating === 'liked' ? '👍 满意' : '👎 不满意'}</div>
                    <div>{details.job.rating.comment || '无评论'}</div>
                  </div>
                )}
              </section>
            </>
          )}
        </main>

        <aside className="admin-stats">
          <h3><BarChart3 size={16} /> 统计信息</h3>
          <div className="stat-card">
            <span>总任务数</span>
            <strong>{stats?.total_jobs ?? 0}</strong>
          </div>
          <div className="stat-card">
            <span>成功率</span>
            <strong>{stats ? `${(stats.success_rate * 100).toFixed(1)}%` : '0%'}</strong>
          </div>
          <div className="stat-card">
            <span>平均耗时</span>
            <strong>{stats?.avg_generation_time_ms ? `${(stats.avg_generation_time_ms / 1000).toFixed(1)}s` : '--'}</strong>
          </div>
          <div className="stat-card">
            <span>点赞 / 点踩</span>
            <strong>{stats?.by_rating?.liked || 0} / {stats?.by_rating?.disliked || 0}</strong>
          </div>
        </aside>
      </div>
    </div>
  );
}
