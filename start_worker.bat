@echo off
echo ==============================================
echo   DOC-AI Background Worker (Huey)
echo ==============================================
echo.
echo Starting task queue consumer...
echo This window must stay open to process reviews.
echo.

huey_consumer.py app.huey_queue

pause
