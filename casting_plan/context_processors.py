from django.db import connection

def notifications(request):
    if request.user.is_authenticated:
        try:
            # Check defensively if migration has run and table exists
            tables = connection.introspection.table_names()
            if 'casting_plan_notification' in tables:
                from .models import Notification
                items = Notification.objects.all()[:5]
                unread_count = Notification.objects.filter(is_read=False).count()
                return {
                    'notifications': items,
                    'unread_notifications_count': unread_count
                }
        except Exception:
            pass
            
    return {
        'notifications': [],
        'unread_notifications_count': 0
    }
