import { LucideIcon } from "lucide-react";

interface FeatureCardProps {
    icon: LucideIcon;
    title: string;
    onClick: () => void;
    onMouseEnter?: () => void;
    onMouseLeave?: () => void;
    disabled?: boolean;
}

export default function FeatureCard({
    icon: Icon,
    title,
    onClick,
    onMouseEnter,
    onMouseLeave,
    disabled = false,
}: FeatureCardProps) {
    return (
        <button
            onClick={onClick}
            onMouseEnter={onMouseEnter}
            onMouseLeave={onMouseLeave}
            disabled={disabled}
            className="group w-full flex items-center gap-4 bg-slate-800/50 hover:bg-slate-700/80 border border-slate-700/50 hover:border-slate-600 rounded-xl p-4 transition-all duration-200 text-left disabled:opacity-50 disabled:cursor-not-allowed"
        >
            {/* Icon Area */}
            <div className="flex-shrink-0 p-2.5 bg-slate-700/50 rounded-lg group-hover:bg-indigo-500/20 group-hover:text-indigo-300 transition-colors">
                <Icon size={20} className="text-slate-400 group-hover:text-indigo-400" />
            </div>

            {/* Content Area */}
            <div className="flex-1 min-w-0">
                <h3 className="text-base font-display text-slate-200 group-hover:text-white transition-colors">{title}</h3>
            </div>

            {/* Chevron or indicator could be nice, but keeping it simple as requested: [Icon] [Title] */}
        </button>
    );
}
