'use client';

export type CheckStatus = 'not_required' | 'required_check' | 'checked_ok' | 'checked_rejected';

interface CommentStatusBadgeProps {
  status: CheckStatus | string;
  showLabel?: boolean;
  size?: 'sm' | 'md';
}

const statusConfig: Record<CheckStatus, {
  label: string;
  icon: string;
  bg: string;
  text: string;
  border: string;
}> = {
  not_required: {
    label: 'No Check Needed',
    icon: 'check_circle',
    bg: 'bg-gray-100 dark:bg-gray-800',
    text: 'text-gray-500 dark:text-gray-400',
    border: 'border-gray-200 dark:border-gray-700'
  },
  required_check: {
    label: 'Needs Review',
    icon: 'pending',
    bg: 'bg-amber-100 dark:bg-amber-900/30',
    text: 'text-amber-700 dark:text-amber-400',
    border: 'border-amber-200 dark:border-amber-800'
  },
  checked_ok: {
    label: 'Approved',
    icon: 'verified',
    bg: 'bg-emerald-100 dark:bg-emerald-900/30',
    text: 'text-emerald-700 dark:text-emerald-400',
    border: 'border-emerald-200 dark:border-emerald-800'
  },
  checked_rejected: {
    label: 'Changes Needed',
    icon: 'error',
    bg: 'bg-red-100 dark:bg-red-900/30',
    text: 'text-red-700 dark:text-red-400',
    border: 'border-red-200 dark:border-red-800'
  }
};

export default function CommentStatusBadge({
  status,
  showLabel = true,
  size = 'sm'
}: CommentStatusBadgeProps) {
  const normalizedStatus = (status || 'not_required') as CheckStatus;
  const config = statusConfig[normalizedStatus] || statusConfig.not_required;

  const sizeClasses = size === 'sm'
    ? 'text-xs px-2 py-0.5 gap-1'
    : 'text-sm px-3 py-1 gap-1.5';

  const iconSize = size === 'sm' ? 'text-sm' : 'text-base';

  return (
    <span
      className={`inline-flex items-center rounded-full border font-medium ${config.bg} ${config.text} ${config.border} ${sizeClasses}`}
      title={config.label}
    >
      <span className={`material-symbols-outlined ${iconSize}`}>{config.icon}</span>
      {showLabel && <span>{config.label}</span>}
    </span>
  );
}
