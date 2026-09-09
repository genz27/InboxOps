import * as React from 'react';
import { cn } from '@/src/lib/utils';

export interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: 'default' | 'destructive' | 'outline' | 'secondary' | 'ghost' | 'link';
  size?: 'default' | 'sm' | 'lg' | 'icon';
}

const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant = 'default', size = 'default', ...props }, ref) => {
    return (
      <button
        ref={ref}
        className={cn(
          'inline-flex items-center justify-center whitespace-nowrap rounded-md text-sm font-medium transition duration-200 ease-[cubic-bezier(0.32,0.72,0,1)] focus-visible:outline-none disabled:pointer-events-none disabled:opacity-40 active:scale-[0.98]',
          {
            'bg-white text-black hover:bg-white/90': variant === 'default',
            'bg-[#ff4d4d] text-white hover:bg-[#ff4d4d]/90': variant === 'destructive',
            'border border-white/15 bg-transparent text-white hover:bg-white/5': variant === 'outline',
            'bg-white/10 text-white hover:bg-white/15': variant === 'secondary',
            'text-white/70 hover:bg-white/5 hover:text-white': variant === 'ghost',
            'text-white underline-offset-4 hover:underline': variant === 'link',
            'h-9 px-3.5': size === 'default',
            'h-8 rounded-md px-2.5 text-xs': size === 'sm',
            'h-10 rounded-md px-8': size === 'lg',
            'h-8 w-8': size === 'icon',
          },
          className
        )}
        {...props}
      />
    );
  }
);
Button.displayName = 'Button';

export { Button };
