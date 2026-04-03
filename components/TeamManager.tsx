import React, { useEffect, useState } from "react";
import { UserPlus, Loader2 } from "lucide-react";
import { useRole } from "../contexts/RoleContext";
import { useAppAuth } from "../contexts/AppAuthContext";
import { getGraphAccessToken } from "../lib/graph";
import { getCleanTrackUsers, upsertUser } from "../repositories/usersRepo";
import { AppSelect } from "./ui";

type RoleOption = "Admin" | "Manager" | "Accounts";

interface TeamRow {
  id: string;
  fullName: string;
  email: string;
  role: RoleOption;
  active: boolean;
}

const TeamManager: React.FC = () => {
  const { isAdmin } = useRole();
  const { user } = useAppAuth();

  const [rows, setRows] = useState<TeamRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const [formName, setFormName] = useState("");
  const [formEmail, setFormEmail] = useState("");
  const [formRole, setFormRole] = useState<RoleOption>("Manager");
  const [formActive, setFormActive] = useState(true);
  const [formLoading, setFormLoading] = useState(false);
  const [formError, setFormError] = useState<string | null>(null);

  useEffect(() => {
    if (!isAdmin) {
      setLoading(false);
      return;
    }
    const load = async () => {
      setLoading(true);
      setError(null);
      const token = await getGraphAccessToken();
      if (!token) {
        setError("Sign in with Microsoft to manage team profiles.");
        setLoading(false);
        return;
      }
      try {
        const users = await getCleanTrackUsers(token);
        setRows(
          users.map((u) => ({
            id: u.id,
            fullName: u.fullName,
            email: u.email,
            role: (u.role === "Admin" ? "Admin" : u.role === "Accounts" ? "Accounts" : "Manager") as RoleOption,
            active: u.active,
          }))
        );
      } catch (e) {
        const msg = e instanceof Error ? e.message : "Failed to load team profiles.";
        setError(msg);
      } finally {
        setLoading(false);
      }
    };
    load();
  }, [isAdmin]);

  const handleUpsert = async (payload: { fullName: string; email: string; role: RoleOption; active: boolean }) => {
    const token = await getGraphAccessToken();
    if (!token) {
      setFormError("Sign in with Microsoft to save team profiles.");
      return;
    }
    setFormLoading(true);
    setFormError(null);
    try {
      await upsertUser(token, {
        fullName: payload.fullName,
        email: payload.email,
        role: payload.role,
        active: payload.active,
      });
      const users = await getCleanTrackUsers(token);
      setRows(
        users.map((u) => ({
          id: u.id,
          fullName: u.fullName,
          email: u.email,
          role: (u.role === "Admin" ? "Admin" : u.role === "Accounts" ? "Accounts" : "Manager") as RoleOption,
          active: u.active,
        }))
      );
      setFormName("");
      setFormEmail("");
      setFormRole("Manager");
      setFormActive(true);
    } catch (e) {
      const msg = e instanceof Error ? e.message : "Failed to save profile.";
      setFormError(msg);
    } finally {
      setFormLoading(false);
    }
  };

  if (!isAdmin) {
    return (
      <div className="p-6 text-sm text-gray-500">
        Only Admins can manage team profiles.
      </div>
    );
  }

  return (
    <div className="space-y-6 sm:space-y-8 animate-fadeIn">
      <div className="flex flex-col sm:flex-row sm:justify-between sm:items-end gap-4 border-b border-[#edeef0] pb-4">
        <div className="min-w-0">
          <h2 className="text-2xl sm:text-3xl font-bold text-gray-900">Team</h2>
          <p className="text-gray-500 text-sm mt-1">
            Manage Admin, Manager, and Accounts profiles used across Sites, Timesheets, and permissions.
          </p>
        </div>
      </div>

      {error && (
        <div className="bg-amber-50 border border-amber-200 text-amber-800 px-4 py-3 rounded-lg text-sm">
          {error}
        </div>
      )}

      <div className="bg-white border border-[#edeef0] rounded-lg shadow-sm p-4 space-y-3 max-w-3xl">
        <div className="flex items-center gap-2 mb-1">
          <UserPlus size={16} className="text-gray-700" />
          <p className="text-xs font-bold text-gray-500 uppercase tracking-widest">
            Add team member
          </p>
        </div>
        <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
          <div className="md:col-span-2">
            <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
              Full name
            </label>
            <input
              type="text"
              value={formName}
              onChange={(e) => setFormName(e.target.value)}
              className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm"
              placeholder="e.g. Jane Manager"
              disabled={formLoading}
            />
          </div>
          <div className="md:col-span-2">
            <label className="block text-[10px] font-bold text-gray-400 uppercase mb-1">
              Email *
            </label>
            <input
              type="email"
              value={formEmail}
              onChange={(e) => setFormEmail(e.target.value)}
              className="w-full border border-[#edeef0] rounded-lg px-2 py-1.5 text-sm"
              placeholder="manager@company.com"
              disabled={formLoading}
            />
          </div>
        </div>
        <div className="flex flex-wrap items-end gap-3">
          <div className="min-w-[10rem]">
            <AppSelect
              label="Role"
              value={formRole}
              onChange={(v) => setFormRole(v as RoleOption)}
              options={[
                { value: "Manager", label: "Manager" },
                { value: "Accounts", label: "Accounts / Financial Controller" },
                { value: "Admin", label: "Admin" },
              ]}
              disabled={formLoading}
              fullWidth
              triggerClassName="rounded-xl px-4 py-2.5 text-sm font-semibold shadow-sm"
            />
          </div>
          <label className="flex items-center gap-2 text-sm text-gray-700 pb-2.5">
            <input
              type="checkbox"
              checked={formActive}
              onChange={(e) => setFormActive(e.target.checked)}
              className="rounded border-gray-300"
              disabled={formLoading}
            />
            Active
          </label>
          <button
            type="button"
            onClick={() => {
              if (!formEmail.trim()) {
                setFormError("Email is required.");
                return;
              }
              handleUpsert({
                fullName: formName.trim() || formEmail.trim(),
                email: formEmail.trim(),
                role: formRole,
                active: formActive,
              });
            }}
            disabled={formLoading}
            className="inline-flex items-center justify-center gap-1.5 px-4 py-2.5 rounded-xl text-[11px] font-bold uppercase tracking-wide so-btn-primary shadow-sm transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
          >
            {formLoading && <Loader2 className="animate-spin" size={12} />}
            Save profile
          </button>
        </div>
        {formError && (
          <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded px-2 py-1 mt-1">
            {formError}
          </p>
        )}
      </div>

      <div className="border border-[#edeef0] rounded-lg bg-white shadow-sm overflow-hidden">
        <table className="w-full border-collapse text-left table-fixed">
          <colgroup>
            <col style={{ width: '28%' }} />
            <col style={{ width: '32%' }} />
            <col style={{ width: '20%' }} />
            <col style={{ width: '20%' }} />
          </colgroup>
          <thead className="bg-[#fcfcfb] border-b border-[#edeef0]">
            <tr className="text-[9px] font-bold text-gray-500 uppercase tracking-widest">
              <th className="py-1.5 px-1.5">Name</th>
              <th className="py-1.5 px-1.5">Email</th>
              <th className="py-1.5 px-1.5">Role</th>
              <th className="py-1.5 px-1.5">Active</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr>
                <td colSpan={4} className="py-6 px-1.5 text-[11px] text-gray-500">
                  <div className="flex items-center gap-2">
                    <Loader2 className="animate-spin" size={16} /> Loading team…
                  </div>
                </td>
              </tr>
            ) : rows.length === 0 ? (
              <tr>
                <td colSpan={4} className="py-6 px-1.5 text-[11px] text-gray-500">
                  No team members yet. Add a profile above.
                </td>
              </tr>
            ) : (
              rows.map((row) => (
                <tr key={row.id} className="border-b border-[#edeef0] last:border-b-0 hover:bg-[#f7f6f3] transition-colors">
                  <td className="py-1.5 px-1.5 text-[11px] text-gray-900">{row.fullName}</td>
                  <td className="py-1.5 px-1.5 text-[11px] text-gray-700 break-words">{row.email}</td>
                  <td className="py-1.5 px-1.5 min-w-0">
                    <AppSelect
                      id={`team-role-${row.id}`}
                      value={row.role}
                      onChange={(v) =>
                        handleUpsert({
                          fullName: row.fullName,
                          email: row.email,
                          role: v as RoleOption,
                          active: row.active,
                        })
                      }
                      options={[
                        { value: "Manager", label: "Manager" },
                        { value: "Accounts", label: "Accounts / Financial Controller" },
                        { value: "Admin", label: "Admin" },
                      ]}
                      disabled={formLoading}
                      size="sm"
                      triggerClassName="text-[11px] min-h-0 py-1"
                    />
                  </td>
                  <td className="py-1.5 px-1.5">
                    <label className="inline-flex items-center gap-1 text-[11px] text-gray-700">
                      <input
                        type="checkbox"
                        checked={row.active}
                        onChange={(e) =>
                          handleUpsert({
                            fullName: row.fullName,
                            email: row.email,
                            role: row.role,
                            active: e.target.checked,
                          })
                        }
                        className="rounded border-gray-300"
                        disabled={formLoading}
                      />
                      <span>{row.active ? "Active" : "Inactive"}</span>
                    </label>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};

export default TeamManager;

