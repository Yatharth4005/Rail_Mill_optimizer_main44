from django.apps import AppConfig
import os
import shutil

class CastingPlanConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'casting_plan'

    def ready(self):
        from django.conf import settings
        
        # 1. Copy custom logo & artwork files
        dest_dir = os.path.join(settings.BASE_DIR, 'casting_plan', 'static', 'casting_plan')
        os.makedirs(dest_dir, exist_ok=True)

        logo_src = os.path.join(settings.BASE_DIR, 'logo-for-dark-dark.png')
        if os.path.exists(logo_src):
            try:
                shutil.copy(logo_src, os.path.join(dest_dir, 'logo.png'))
            except Exception:
                pass

        art_mappings = {
            r"C:\Users\kaush\.gemini\antigravity-ide\brain\ae6dc8bc-8958-4884-8520-6fad7dec593d\hero_banner_panoramic_1789474352294.png": "hero-banner-art.png",
            r"C:\Users\kaush\.gemini\antigravity-ide\brain\ae6dc8bc-8958-4884-8520-6fad7dec593d\rail_mill_section_flow_1789474076683.png": "dashboard-flow-art.png",
            r"C:\Users\kaush\.gemini\antigravity-ide\brain\ae6dc8bc-8958-4884-8520-6fad7dec593d\railway_sunset_tracks_1789474110793.png": "target-bullseye-art.png",
        }
        for art_src, art_filename in art_mappings.items():
            if os.path.exists(art_src):
                try:
                    shutil.copy(art_src, os.path.join(dest_dir, art_filename))
                except Exception:
                    pass

        # 2. Database auto-initialization (defensive check for migrations)
        from django.db import connection
        try:
            tables = connection.introspection.table_names()
            
            # Check auth_user table for admin credentials
            if 'auth_user' in tables:
                from django.contrib.auth.models import User
                if not User.objects.filter(username='admin').exists():
                    User.objects.create_superuser('admin', 'admin@example.com', 'admin123')
                    
            # Check notification table for initial data
            if 'casting_plan_notification' in tables:
                from casting_plan.models import Notification
                if Notification.objects.count() == 0:
                    Notification.objects.create(
                        title="System Ready",
                        message="The Rail Mill Optimizer portal has been successfully initialized with rules and templates.",
                        is_read=False
                    )
                    Notification.objects.create(
                        title="New Rules Imported",
                        message="Imported 52 Section Rules and 66 LFA Batch Rules from Rail Mill workbooks.",
                        is_read=False
                    )
                    Notification.objects.create(
                        title="Optimization Opportunity",
                        message="Grade mapping configured for external grades. Run the optimizer to see RM yield gains.",
                        is_read=False
                    )
        except Exception:
            # Silence issues if database is not yet migrated/accessible
            pass
