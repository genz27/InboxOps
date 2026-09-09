import * as React from 'react';
import { cn } from '@/src/lib/utils';

export type BadgeProps = React.ComponentPropsWithoutRef<'div'> & {
  variant?: 'default' | 'secondary' | 'destructive' | 'outline';
};

const Badge = React.forwardRef<HTMLDivElement, BadgeProps>(
  ({ className, variant = 'default', ...props }, ref) => {
    return (
      <div
        ref={ref}
        className={cn(
          'inline-flex items-center rounded-md border px-2 py-0.5 text-[11px] font-medium',
          {
            'border-transparent bg-white text-black': variant === 'default',
            'border-transparent bg-white/10 text-white/80': variant === 'secondary',
            'border-transparent bg-[#ff4d4d]/15 text-[#ff8080]': variant === 'destructive',
            'border-white/15 text-white/70': variant === 'outline',
          },
          className
        )}
        {...props}
      />
    );
  }
);
Badge.displayName = 'Badge';

export { Badge };
